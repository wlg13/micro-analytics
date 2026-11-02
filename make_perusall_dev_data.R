# Generate fake Perusall exports for development, so Perusall code is never
# written or tested against real student data. Mirrors the headers and value
# formats of the real exports (all-data export + gradebook) as described by
# the instructor; nothing here is read from the real files.
#
# Writes to $MA/dev/<term>/perusall/ and refuses to run unless MA_MODE=dev.
# Emails in users.csv and the gradebook are taken from the dev (hashed) merged
# roster, so joins to the roster can be tested; a few deliberately don't match.
#
# Usage (from the project folder): MA_MODE=dev Rscript make_perusall_dev_data.R
#   Run merge_rosters.R in dev mode first.
#
# GUESSES (no example seen yet; fix here if the real files differ):
#   - User ID, Document ID, Assignment ID, Comment ID: 17-character IDs like Student ID
#   - Course role: "student" / "instructor"
#   - Content to assign, Document, Document name: the reading's name
#   - Submission deadline, Opened at, Closed at: same M/D/YYYY H:MM format as Created
#   - Group set name, Part, Range, Starred by me?, Total viewing time (minutes)
#   - Average score: mean of the assignment scores, 2 decimals

library(tidyverse)

source("ma_paths.R")
dirs <- ma_dirs()
if (dirs$mode != "dev") stop("make_perusall_dev_data.R only writes fake data in dev mode; set MA_MODE=dev.")

out_dir <- file.path(dirs$inputs, "perusall")
if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)

roster_path <- file.path(dirs$outputs, "merged_roster.csv")
if (!file.exists(roster_path)) stop("Run merge_rosters.R in dev mode first: ", roster_path, " not found.")
roster_emails <- read_csv(roster_path, col_types = cols(.default = "c"))$Email

set.seed(104)
n_students <- 20

# Perusall-style IDs: 17 characters, no easily confused characters
id_chars <- strsplit("23456789ABCDEFGHJKLMNPQRSTWXYZabcdefghijkmnopqrstuvwxyz", "")[[1]]
new_ids <- function(n) replicate(n, paste(sample(id_chars, 17, replace = TRUE), collapse = ""))
fake_hash <- function(n) replicate(n, paste(sample(c(0:9, letters[1:6]), 16, replace = TRUE), collapse = ""))

# M/D/YYYY H:MM, 24-hour, no leading zeros (e.g. 8/23/2026 19:42)
fmt_time <- function(t) {
  out <- paste0(month(t), "/", mday(t), "/", year(t), " ", hour(t), ":", sprintf("%02d", minute(t)))
  if_else(is.na(t), "", out)
}
et <- function(x) as.POSIXct(x, tz = "America/New_York")

first_names <- c("Alex", "Jordan", "Taylor", "Morgan", "Casey", "Riley", "Avery", "Quinn",
                 "Jamie", "Drew", "Reese", "Skyler", "Parker", "Rowan", "Emerson", "Hayden")
last_names <- c("Testperson", "Sample", "Example", "Placeholder", "Fakename", "Mockley",
                "Dummy", "Synthetic", "Madeup", "Fictional", "Notreal", "Specimen")

# 17 students on the dev roster, 3 not (e.g. personal email or not enrolled)
students <- tibble(
  user_id = new_ids(n_students),
  student_id = new_ids(n_students),
  first = sample(first_names, n_students, replace = TRUE),
  last = sample(last_names, n_students, replace = TRUE),
  email = c(sample(roster_emails, n_students - 3), fake_hash(3)),
  group = sample(c(6, 15, 22, 23), n_students, replace = TRUE),
  effort = runif(n_students, 0.2, 1)  # drives how much each student annotates and reads
)
instructor <- tibble(user_id = new_ids(1), student_id = new_ids(1), first = "Pat", last = "Instructor",
                     email = fake_hash(1))

# Reading assignment names and deadlines as in the real gradebook (course
# material, not student data); the third name contains a comma on purpose.
readings <- tibble(
  assignment_id = new_ids(3),
  document_id = new_ids(3),
  name = c("GDP 1: Total Production and Does it Measure What We Want It To (Ch. 8.1 & 8.2)",
           "Stocks and Bonds (Ch. 6.2)",
           "GDP 2: Real GDP and, and Other Measures of Total Production and Income (Ch. 8.3 & 8.4)"),
  document_name = c("Ch. 8.1 & 8.2", "Ch. 6.2", "Ch. 8.3 & 8.4"),
  opens = et(c("2026-08-24 08:00", "2026-08-24 08:00", "2026-08-31 08:00")),
  due = et(c("2026-08-31 23:59", "2026-08-31 23:59", "2026-09-07 23:59")),
  due_text = c("Due August 31, 2026 at 11:59 PM EDT", "Due August 31, 2026 at 11:59 PM EDT",
               "Due September 7, 2026 at 11:59 PM EDT"),
  created = et(c("2026-08-20 10:15", "2026-08-20 10:22", "2026-08-21 09:40"))
)

random_time <- function(from, to) from + runif(length(from)) * as.numeric(difftime(to, from, units = "secs"))

# ---- annotations: top-level comments, then replies to some of them ----
phrases <- c("I think this means", "Why does", "This connects to the lecture on", "I am confused about",
             "An example of this is", "Does this imply that", "I agree with the point about",
             "How would this change if")
topics <- c("real GDP", "inflation", "bond prices", "interest rates", "household production",
            "the price level", "stock returns", "nominal values")
fake_comment <- function(n) {
  map_chr(seq_len(n), function(i) {
    filler <- paste(sample(c("the", "text", "says", "that", "this", "is", "important", "because",
                             "it", "affects", "how", "we", "measure", "output"), sample(4:40, 1), replace = TRUE),
                    collapse = " ")
    paste(sample(phrases, 1), sample(topics, 1), filler)
  })
}

top <- students %>%
  crossing(readings %>% select(document_name, opens, due)) %>%
  mutate(n = rpois(n(), 3 * effort) * (runif(n()) > 0.15)) %>%  # some students skip a reading
  uncount(n) %>%
  mutate(parent = "", created = random_time(opens, due))

replies <- top %>%
  slice_sample(prop = 0.3) %>%
  transmute(document_name, parent_created = created, due)
replies <- replies %>%
  bind_cols(slice_sample(students, n = nrow(replies), replace = TRUE)) %>%
  mutate(created = random_time(parent_created, due))

annotations <- bind_rows(top, replies %>% select(-parent_created)) %>%
  mutate(comment_id = new_ids(n()))
# Each reply points at a top-level comment on the same document
top_ids <- annotations %>% filter(row_number() <= nrow(top)) %>% select(comment_id, document_name)
reply_rows <- which(seq_len(nrow(annotations)) > nrow(top))
annotations$parent[reply_rows] <- map_chr(reply_rows, function(i) {
  sample(top_ids$comment_id[top_ids$document_name == annotations$document_name[i]], 1)
})

annotations <- annotations %>%
  mutate(
    text = fake_comment(n()),
    highlighted = if_else(parent == "", paste("Highlighted passage about", sample(topics, n(), replace = TRUE)), ""),
    edited = if_else(runif(n()) < 0.1, created + runif(n(), 60, 3600), created),
    page = sample(1:30, n(), replace = TRUE),
    start = sample(100:5000, n(), replace = TRUE)
  ) %>%
  arrange(created)

annotations_out <- annotations %>%
  transmute(
    `Comment ID` = comment_id,
    `In response to comment ID` = parent,
    `Last name` = last,
    `First name` = first,
    `Student ID` = student_id,
    `Group set name` = "Discussion groups",
    `Group number` = group,
    `Highlighted text` = highlighted,
    `Comment text` = text,
    `Word count` = str_count(text, "\\S+"),
    Type = sample(c("comment", "question"), n(), replace = TRUE, prob = c(24163, 12682)),
    Score = sample(c("low", "medium", "high"), n(), replace = TRUE, prob = c(11, 10, 5)),
    Created = fmt_time(created),
    `Last edited at` = fmt_time(edited),
    Replies = map_int(comment_id, ~ sum(annotations$parent == .x)),
    Upvoters = sample(0:3, n(), replace = TRUE, prob = c(0.8, 0.12, 0.05, 0.03)),
    `Starred by me?` = "",
    Status = "",
    Document = document_name,
    Part = "",
    `Page number` = page,
    Range = paste0(start, "-", start + sample(20:300, n(), replace = TRUE))
  )

# ---- document-views: reading sessions; many have no active engagement time ----
document_views <- students %>%
  crossing(readings %>% select(document_id, assignment_id, opens, due)) %>%
  mutate(n = rpois(n(), 4 * effort)) %>%
  uncount(n) %>%
  mutate(
    opened = random_time(opens, due + 3 * 86400),  # some open after the deadline
    closed = opened + runif(n(), 30, 5400),
    active = if_else(runif(n()) < 0.4, NA_real_, 60 * pmax(1, round(runif(n(), 0, 1) * as.numeric(closed - opened, units = "mins"))))
  ) %>%
  arrange(opened) %>%
  transmute(
    `Document ID` = document_id,
    `Assignment ID` = assignment_id,
    `User ID` = user_id,
    `Opened at` = fmt_time(opened),
    `Closed at` = fmt_time(closed),
    `Active engagement time` = active
  )

# ---- users ----
active_min <- round(runif(n_students, 37, 855) * students$effort, 2)
active_min <- if_else(runif(n_students) < 0.3, round(active_min), active_min)
users <- bind_rows(
  tibble(`User ID` = instructor$user_id, `Last name` = instructor$last, `First name` = instructor$first,
         `Student ID` = instructor$student_id, `Email address` = instructor$email, `Course role` = "instructor",
         `Total viewing time (minutes)` = NA_real_, `Total active engagement time (minutes)` = NA_real_),
  tibble(`User ID` = students$user_id, `Last name` = students$last, `First name` = students$first,
         `Student ID` = students$student_id, `Email address` = students$email, `Course role` = "student",
         `Total viewing time (minutes)` = round(active_min * runif(n_students, 1.2, 2.5), 2),
         `Total active engagement time (minutes)` = active_min)
)

documents <- readings %>%
  transmute(`Document ID` = document_id, `Document name` = document_name,
            `Uploaded by user ID` = instructor$user_id, Created = fmt_time(created))

assignments <- readings %>%
  transmute(`Content to assign` = document_name, `Assignment name` = name,
            `Submission deadline` = fmt_time(due), Created = fmt_time(created))

# ---- gradebook: header, then a due-date line and a points line, then students ----
# Score 0-3 per assignment from the student's annotations on it; 0 if none.
scores <- students %>%
  crossing(readings %>% select(document_name, name)) %>%
  left_join(count(annotations, student_id, document_name), by = c("student_id", "document_name")) %>%
  mutate(score = pmin(3L, coalesce(n, 0L))) %>%
  select(student_id, name, score) %>%
  pivot_wider(names_from = name, values_from = score)

gradebook <- students %>%
  left_join(scores, by = "student_id") %>%
  mutate(`Average score` = round(rowMeans(across(all_of(readings$name))), 2)) %>%
  arrange(last, first) %>%
  transmute(`Last name` = last, `First name` = first, `Student ID` = student_id,
            Group = paste("Group", group), Email = email, `Average score` = as.character(`Average score`),
            across(all_of(readings$name), as.character))
info_rows <- tibble(
  `Last name` = "", `First name` = "", `Student ID` = "", Group = "", Email = "",
  `Average score` = c("Released & past deadline", "0-3 points")
)
for (i in seq_len(nrow(readings))) info_rows[[readings$name[i]]] <- c(readings$due_text[i], "0-3 points")
gradebook <- bind_rows(info_rows, gradebook)

write_csv(users, file.path(out_dir, "users.csv"), na = "")
write_csv(annotations_out, file.path(out_dir, "annotations.csv"), na = "")
write_csv(assignments, file.path(out_dir, "assignments.csv"), na = "")
write_csv(documents, file.path(out_dir, "documents.csv"), na = "")
write_csv(document_views, file.path(out_dir, "document-views.csv"), na = "")
write_csv(gradebook, file.path(out_dir, "econ-104-fa26-gradebook.csv"), na = "")

message("Wrote 6 fake Perusall files to ", out_dir, ": ", nrow(users) - 1, " students, ",
        nrow(annotations_out), " annotations, ", nrow(document_views), " document views")
