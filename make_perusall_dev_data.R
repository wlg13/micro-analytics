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
# Real formats this mimics:
#   - All readings are sections of ONE epub document (the textbook); documents.csv
#     also lists a peer-review upload. Readings are told apart by Assignment ID
#     (document-views) and by the epub chapter index in Range / Section.
#   - annotations, assignments, documents: Eastern time, M/D/YYYY H:MM (e.g. 8/23/2026 19:42)
#   - document-views (and chapter-views): UTC ISO 8601 with milliseconds (2026-08-27T19:04:46.263Z)
#   - Active engagement time: seconds in multiples of 60, often blank
#   - Gradebook: scores 0-3, missed = 0; Average score = unrounded mean over the
#     assignments past their deadline at download time (8 of 9 on 10/26)
# GUESSES (no example seen yet; fix here if the real files differ):
#   - the annotations Document column holds the document's name
#   - the epub chapter ranges of each reading (made up)

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
download <- as.POSIXct("2026-10-26 21:34", tz = "America/New_York")

# Perusall-style IDs: 17 characters, no easily confused characters
id_chars <- strsplit("23456789ABCDEFGHJKLMNPQRSTWXYZabcdefghijkmnopqrstuvwxyz", "")[[1]]
new_ids <- function(n) replicate(n, paste(sample(id_chars, 17, replace = TRUE), collapse = ""))
fake_hash <- function(n) replicate(n, paste(sample(c(0:9, letters[1:6]), 16, replace = TRUE), collapse = ""))

# Eastern M/D/YYYY H:MM, 24-hour, no leading zeros (e.g. 8/23/2026 19:42)
fmt_et <- function(t) {
  t <- with_tz(t, "America/New_York")
  out <- paste0(month(t), "/", mday(t), "/", year(t), " ", hour(t), ":", sprintf("%02d", minute(t)))
  if_else(is.na(t), "", out)
}
# UTC ISO 8601 with milliseconds (e.g. 2026-08-27T19:04:46.263Z)
fmt_utc <- function(t) {
  t <- with_tz(t, "UTC")
  paste0(format(t, "%Y-%m-%dT%H:%M:"), sprintf("%06.3f", floor(second(t) * 1000) / 1000), "Z")
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

documents <- tibble(
  `Document ID` = new_ids(2),
  `Document name` = c("Peer review submission upload", "Macroeconomics, 9/e"),
  `Uploaded by user ID` = "",
  Created = c("8/12/2024 9:03", "3/28/2024 10:57")
)
textbook_id <- documents$`Document ID`[2]
textbook_name <- documents$`Document name`[2]

# Reading assignments and deadlines as in the real assignments.csv (course
# material, not student data); one name contains a comma on purpose.
# ch_from/ch_to: made-up epub chapter indexes of each reading.
readings <- tribble(
  ~name, ~due, ~ch_from, ~ch_to,
  "GDP 1: Total Production and Does it Measure What We Want It To (Ch. 8.1 & 8.2)", "2026-08-31 23:59", 187, 193,
  "Stocks and Bonds (Ch. 6.2)", "2026-08-31 23:59", 142, 146,
  "GDP 2: Real GDP and, and Other Measures of Total Production and Income (Ch. 8.3 & 8.4)", "2026-09-07 23:59", 194, 202,
  "The Consumer Price Index (CPI) and Its Uses (Ch. 9.4 to 9.7)", "2026-09-14 23:59", 220, 226,
  "Long-Run Growth 1 (Ch. 10.1)", "2026-09-30 23:59", 238, 241,
  "Long-Run Growth 2 (Ch. 11.1 to 11.3)", "2026-10-07 23:59", 257, 269,
  "Employment and Unemployment (Ch. 9.1 & 9.2)", "2026-10-26 23:59", 212, 218,
  "The Business Cycle (Ch. 10.3)", "2026-10-26 23:59", 245, 248,
  "Aggregate Demand and Supply (Ch. 13.1 & 13.2)", "2026-11-09 23:59", 320, 327
) %>%
  mutate(
    due = et(due),
    assignment_id = new_ids(n()),
    opens = lag(due, default = et("2026-08-22 12:00")) - 3 * 86400,
    until = pmin(due, download),  # activity stops at the download
    due_text = paste0("Due ", format(due, "%B "), mday(due), format(due, ", %Y at 11:59 PM "),
                      if_else(dst(due), "EDT", "EST"))
  )

# Random times between from and to, weighted toward the evening (Eastern)
random_time <- function(from, to) {
  t <- from + runif(length(from)) * as.numeric(difftime(to, from, units = "secs"))
  late <- runif(length(t)) < 0.6
  day <- floor_date(with_tz(t, "America/New_York"), "day")
  t[late] <- pmin(to[late], pmax(from[late], day[late] + runif(sum(late), 18, 24) * 3600))
  t
}

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

# Fewer comments on the reading that is not yet due at the download
top <- students %>%
  crossing(readings %>% select(name, opens, until, ch_from, ch_to)) %>%
  mutate(n = rpois(n(), if_else(until < download, 2, 0.2) * effort) * (runif(n()) > 0.15)) %>%
  uncount(n) %>%
  mutate(comment_id = new_ids(n()), parent = "", created = random_time(opens, until),
         chapter = map2_int(ch_from, ch_to, ~ sample(.x:.y, 1L)))

# Each reply points at a top-level comment on the same reading and comes after it
replies <- top %>%
  filter(created < until - 120) %>%
  slice_sample(prop = 0.3) %>%
  transmute(name, parent = comment_id, parent_created = created, until, chapter)
replies <- replies %>%
  bind_cols(slice_sample(students, n = nrow(replies), replace = TRUE)) %>%
  mutate(comment_id = new_ids(n()),
         created = random_time(parent_created + 60, until))

annotations <- bind_rows(top, replies %>% select(-parent_created)) %>%
  mutate(
    text = fake_comment(n()),
    highlighted = if_else(parent == "", paste("Highlighted passage about", sample(topics, n(), replace = TRUE)), ""),
    edited = if_else(runif(n()) < 0.1, created + runif(n(), 60, 3600), created),
    start = sample(0:400, n(), replace = TRUE),
    range = sprintf('{"cfi":"epubcfi(/8/%d[P7001018354000000000000000003E7A]!/4/2/2/4/1:%d,1:%d)","chapter":%d,"end":%d,"start":%d,"type":"text"}',
                    2 * chapter, start, start + 72, chapter, start + 72, start)
  ) %>%
  arrange(created)

annotations_out <- annotations %>%
  transmute(
    `Comment ID` = comment_id,
    `In response to comment ID` = parent,
    `Last name` = last,
    `First name` = first,
    `Student ID` = student_id,
    `Group set name` = "Group Set 1",
    `Group number` = group,
    `Highlighted text` = highlighted,
    `Comment text` = text,
    `Word count` = str_count(text, "\\S+"),
    Type = sample(c("comment", "question"), n(), replace = TRUE, prob = c(24163, 12682)),
    Score = sample(c("low", "medium", "high"), n(), replace = TRUE, prob = c(11, 10, 5)),
    Created = fmt_et(created),
    `Last edited at` = fmt_et(edited),
    Replies = map_int(comment_id, ~ sum(annotations$parent == .x)),
    Upvoters = sample(0:3, n(), replace = TRUE, prob = c(0.8, 0.12, 0.05, 0.03)),
    `Starred by me?` = "No",
    Status = "",
    Document = textbook_name,
    Part = "",
    `Page number` = "",
    Range = range
  )

# ---- document-views: reading sessions; many very short, many without active time ----
document_views <- students %>%
  crossing(readings %>% select(assignment_id, opens, until)) %>%
  mutate(n = rpois(n(), 3 * effort)) %>%
  uncount(n) %>%
  mutate(
    opened = random_time(opens, until) + runif(n(), 0, 60),
    closed = opened + if_else(runif(n()) < 0.4, runif(n(), 0.1, 1), runif(n(), 30, 5400)),
    secs = as.numeric(difftime(closed, opened, units = "secs")),
    active = if_else(runif(n()) < 0.4 | secs < 60, NA_real_, 60 * pmax(1, round(runif(n(), 0, 1) * secs / 60)))
  ) %>%
  arrange(opened) %>%
  transmute(
    `Document ID` = textbook_id,
    `Assignment ID` = assignment_id,
    `User ID` = user_id,
    `Opened at` = fmt_utc(opened),
    `Closed at` = fmt_utc(closed),
    `Active engagement time` = active
  )

# ---- users ----
active_min <- round(runif(n_students, 37, 855) * students$effort, 2)
active_min <- if_else(runif(n_students) < 0.3, round(active_min), active_min)
users <- bind_rows(
  tibble(`User ID` = instructor$user_id, `Last name` = instructor$last, `First name` = instructor$first,
         `Student ID` = instructor$student_id, `Email address` = instructor$email, `Course role` = "Instructor",
         `Total viewing time (minutes)` = NA_real_, `Total active engagement time (minutes)` = NA_real_),
  tibble(`User ID` = students$user_id, `Last name` = students$last, `First name` = students$first,
         `Student ID` = students$student_id, `Email address` = students$email, `Course role` = "Student",
         `Total viewing time (minutes)` = round(active_min * runif(n_students, 3, 8), 2),
         `Total active engagement time (minutes)` = active_min)
)

assignments <- readings %>%
  transmute(`Content to assign` = "epub", `Assignment name` = name,
            `Submission deadline` = fmt_et(due), Created = "8/11/2026 13:58")

# ---- gradebook: header, then a due-date line and a points line, then students ----
# Score 0-3 per assignment from the student's comments on it; 0 if none.
scores <- students %>%
  crossing(readings %>% select(name)) %>%
  left_join(count(annotations, student_id, name), by = c("student_id", "name")) %>%
  mutate(score = pmin(3L, coalesce(n, 0L))) %>%
  select(student_id, name, score) %>%
  pivot_wider(names_from = name, values_from = score)

# The real gradebook averaged the 8 readings due by 10/26 (scores in eighths),
# so it was downloaded after that day's deadline.
past_due <- readings$name[as_date(readings$due) <= as_date(download)]
gradebook <- students %>%
  left_join(scores, by = "student_id") %>%
  mutate(`Average score` = rowMeans(across(all_of(past_due)))) %>%
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
