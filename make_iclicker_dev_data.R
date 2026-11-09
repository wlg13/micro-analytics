# Generate fake iClicker exports for development, so iClicker code is never
# written or tested against real student data. Writes, for sections 2, 5 and 11:
#   $MA/dev/<term>/iclicker/gradebook/sec_N/iClicker_GradesExport_Reef_MM-DD-YY.csv
#   $MA/dev/<term>/iclicker/class_meetings/sec_N/iClicker_ClassK-Poll.xlsx
# Refuses to run unless MA_MODE=dev.
#
# Emails come from the dev (hashed) merged roster; one student per section is
# not on it (a drop). Hashing hides the @psu.edu domain, so in dev a roster hash
# stands for a good PSU email and the bad cases are literal strings.
#
# Usage (from the project folder): MA_MODE=dev Rscript make_iclicker_dev_data.R
#   Run merge_rosters.R in dev mode first.
#
# Real formats this mimics (confirmed by the instructor, Fa26):
#   Gradebook CSV, one per section, named by iClicker with the download date;
#   same-day downloads get browser suffixes " (1)", " (2)" (section only from the folder).
#     Name ("Last, First", student-typed: some multi-word or lowercase last names),
#     Email, Student ID (student-typed; ignored by the importer), Total = Performance
#     + Participation, Total % ("99.99%"), Class N - Poll newest first (only meetings
#     with polls, here 8..3), points or "No Response".
#   Class-meeting workbook, one per meeting per section:
#     Summary sheet: row 1 "Summary - Class N - Poll", row 2 session date-time
#     ("09/01/26 03:16 PM"), rows 3-4 blank, row 5 group labels "Activity Scores" (B)
#     and "Question Results" (E), row 6 headers Name | Total | Performance | Participation |
#     "Question k<newline>Multiple Choice", row 7 blank, students from row 8 (1 or "NR"), then
#     Possible Points, Average Points, Average % rows directly below.
#     "Question k" sheets: rows 1-12 info (title "Question k - Class N - Poll",
#     date, "Question Type: Multiple Choice", Correct Answers N/A, Responses,
#     No Responses, Possible Points, Average Points, Average %), row 14 Name |
#     Response, students from row 15 (answer letter, "NR" if none); one embedded JPEG of the slide.
#   Any answer earns the point; participation (1 point) is earned by answering >= 75% of a
#   meeting's questions; Class N - Poll = that meeting's Total.
#   Summary Average % has 2 decimals, Question sheets 1.
# Test cases built in (section 11): blank Email, icloud.com Email, bare "abc1234"
#   Email, pau.edu typo -- each with a good Student ID (importer: use it, warn);
#   a drop per section (good email, not on roster: report); a peer-instruction
#   re-vote (same slide twice); class 6 where section 5 got a different Q2.
#   Stop cases (duplicate names, no good email at all) are left to the importer's tests.
# GUESSES (no example seen yet; fix here if the real files differ):
#   - Students who answer nothing get 0s in the Summary; Summary averages are over
#     students who answered something (inferred from one real example)
#   - Summary/Question sheet numbers stored as numbers; percents as text
#   - Slide image placed at D1; the window title shows the section ("Econ 104-11");
#     slide titles "Question k" (a re-vote repeats the first vote's slide)
#   - The line break in "Question k<newline>Multiple Choice": openxlsx saves it as
#     CRLF on Windows; Excel may use LF -- importers should accept either

library(tidyverse)
library(openxlsx)

source("ma_paths.R")
dirs <- ma_dirs()
if (dirs$mode != "dev") stop("make_iclicker_dev_data.R only writes fake data in dev mode; set MA_MODE=dev.")

ic_dir <- file.path(dirs$inputs, "iclicker")
roster_path <- file.path(dirs$outputs, "merged_roster.csv")
if (!file.exists(roster_path)) stop("Run merge_rosters.R in dev mode first: ", roster_path, " not found.")
roster <- read_csv(roster_path, col_types = cols(.default = "c"))

set.seed(104)
sections <- c("2", "5", "11")
n_per_section <- 12
classes <- 3:8                      # meetings with polls (1 and 2 had none)
download <- as.Date("2026-09-18")   # day after class 8

fake_hash <- function(n) replicate(n, paste(sample(c(0:9, letters[1:6]), 16, replace = TRUE), collapse = ""))
fake_user <- function() paste0(paste(sample(letters, 3), collapse = ""), sample(1000:9999, 1))

# Two meetings a week, Tuesday/Thursday, class 1 on Tue 8/25/2026
class_date <- function(n) as.Date("2026-08-25") + (n - 1) %/% 2 * 7 + (n - 1) %% 2 * 2
start_time <- c("2" = "10:35", "5" = "12:05", "11" = "15:15")
fmt_ic <- function(t) format(t, "%m/%d/%y %I:%M %p")   # 09/01/26 03:16 PM

# One CSV line per row, quoting only fields that need it; CRLF line ends.
write_crlf_csv <- function(header, rows, path) {
  quote <- function(v) {
    v <- if_else(is.na(v), "", v)
    if_else(grepl('[",\r\n]', v), paste0('"', gsub('"', '""', v), '"'), v)
  }
  lines <- c(paste(quote(header), collapse = ","),
             apply(as.matrix(rows), 1, function(r) paste(quote(r), collapse = ",")))
  con <- file(path, open = "wb")
  writeLines(enc2utf8(lines), con, sep = "\r\n", useBytes = TRUE)
  close(con)
}

# ---- Students ---------------------------------------------------------------
first_names <- c("Alex", "Jordan", "Taylor", "Morgan", "Casey", "Riley", "Avery", "Quinn",
                 "Jamie", "Drew", "Reese", "Skyler", "Parker", "Rowan", "Emerson", "Hayden")
last_names <- c("Testperson", "Sample", "Example", "Placeholder", "Fakename", "Mockley",
                "Dummy", "Synthetic", "Madeup", "Fictional", "Notreal", "Specimen")

# Unique "Last, First" names, a few typed the way students type them
make_names <- function(n) {
  combos <- crossing(first = first_names, last = last_names) %>% slice_sample(n = n)
  last <- combos$last
  k <- sample(n, 4)
  last[k[1]] <- paste("van", last[k[1]])
  last[k[2]] <- paste("De", last[k[2]])
  last[k[3]] <- paste(last[k[3]], "Placeholder")
  name <- paste0(last, ", ", combos$first)
  name[k[4]] <- tolower(name[k[4]])
  name
}

picked <- roster %>%
  filter(!is.na(Email), Section %in% sections) %>%
  group_by(Section) %>%
  slice_sample(n = n_per_section) %>%
  ungroup() %>%
  transmute(section = Section, email = Email, dropped = FALSE)
students <- bind_rows(picked, tibble(section = sections, email = fake_hash(3), dropped = TRUE)) %>%
  mutate(name = make_names(n()),
         attend = runif(n(), 0.55, 0.97),
         last_class = if_else(dropped, 5L, 99L))   # drops stop after class 5

# Email / Student ID as typed. Most Student IDs are blank or repeat the email
# (sometimes capitalized); a few hold junk. All of that is ignored by the importer.
capital <- function(x) paste0(toupper(substr(x, 1, 1)), substring(x, 2))
students <- students %>%
  mutate(
    gb_email = email,
    gb_sid = map_chr(email, ~ sample(c("", .x, capital(.x), fake_user(), paste0(fake_user(), "@gmail.com")),
                                     1, prob = c(0.6, 0.2, 0.1, 0.05, 0.05)))
  )
# Section 11 warning cases: bad Email, good Student ID
bad <- which(students$section == "11" & !students$dropped)[1:4]
students$gb_email[bad] <- c("", paste0(fake_user(), "@icloud.com"), fake_user(), paste0(fake_user(), "@pau.edu"))
students$gb_sid[bad] <- students$email[bad]

# ---- Questions --------------------------------------------------------------
bank <- tribble(
  ~text,                                                                              ~choices,
  "Which of the following is counted in GDP?",                                        c("A new car", "A used car", "A stock purchase", "A transfer payment"),
  "If nominal GDP rose 5% and prices rose 3%, real GDP rose about",                   c("2%", "3%", "5%", "8%"),
  "A college graduate looking for a first job is",                                    c("Frictionally unemployed", "Structurally unemployed", "Cyclically unemployed", "Not in the labor force"),
  "Which group is NOT in the labor force?",                                           c("Retirees", "Part-time workers", "Job seekers", "Laid-off workers"),
  "How optimistic are you about the economy over the next year?",                     c("Very optimistic", "Somewhat optimistic", "Somewhat pessimistic", "Very pessimistic"),
  "The CPI measures the price of",                                                    c("A fixed basket of goods", "All goods produced", "Stocks and bonds", "Imports only"),
  "If the CPI goes from 200 to 210, the inflation rate is",                           c("5%", "10%", "2%", "210%"),
  "Which bias makes the CPI overstate inflation?",                                    c("Substitution bias", "Sampling bias", "Survivor bias", "Selection bias"),
  "Predict: next month's unemployment rate will be",                                  c("Higher", "About the same", "Lower", "No idea"),
  "Real GDP per person grows mainly because of",                                      c("Productivity", "Population", "Inflation", "Exports"),
  "At 7% growth, real GDP doubles in about",                                          c("10 years", "7 years", "20 years", "70 years"),
  "A bond's price and its interest rate move",                                        c("In opposite directions", "Together", "Independently", "Only in recessions"),
  "Buying a share of stock makes you",                                                c("A part owner", "A lender", "A bondholder", "An employee"),
  "A recession is usually marked by",                                                 c("Falling real GDP", "Rising prices", "Falling interest rates", "Rising exports"),
  "Which is a leading indicator?",                                                    c("Building permits", "Unemployment rate", "CPI", "GDP"),
  "Which do you think matters most for long-run growth?",                             c("Technology", "Education", "Investment", "Institutions"),
  "Discouraged workers are counted as",                                               c("Not in the labor force", "Unemployed", "Employed", "Part-time")
)

# Same questions for every section in a meeting, except class 6 where section 5
# got a different Q2. One re-vote (peer instruction): the same slide asked twice.
plan <- map_dfr(classes, function(n) {
  nq <- sample(4:8, 1)       # real meetings have had 9 questions with a re-vote
  ids <- sample(nrow(bank), nq)
  revote <- rep(NA_integer_, nq)
  if (n %in% c(4, 7)) {      # re-vote of question 2 as question 3
    ids <- append(ids, ids[2], after = 2)
    revote <- append(revote, 2L, after = 2)
  }
  tibble(class = n, q = seq_along(ids), qid = ids, revote_of = revote)
})
plan <- crossing(section = sections, plan) %>%
  mutate(qid = if_else(class == 6 & section == "5" & q == 2,
                       setdiff(seq_len(nrow(bank)), plan$qid[plan$class == 6])[1], qid))

# ---- Slide images -----------------------------------------------------------
img_dir <- file.path(tempdir(), "iclicker_slides")
dir.create(img_dir, showWarnings = FALSE)
make_slide <- function(path, section, k, text, choices) {
  jpeg(path, width = 960, height = 540, quality = 85)
  par(mar = c(0, 0, 0, 0))
  plot.new()
  rect(0, 0.93, 1, 1, col = "grey85", border = NA)
  text(0.01, 0.965, paste0("Econ 104-", section, " - Zoom"), adj = 0, cex = 1.1)
  text(0.05, 0.84, paste("Question", k), adj = 0, cex = 2, font = 2)
  wrapped <- strwrap(text, 55)
  text(0.05, 0.72 - 0.07 * (seq_along(wrapped) - 1), wrapped, adj = 0, cex = 1.6)
  y0 <- 0.72 - 0.07 * length(wrapped) - 0.06
  text(0.08, y0 - 0.09 * (0:3), paste0(LETTERS[1:4], ". ", choices), adj = 0, cex = 1.5)
  dev.off()
}

# ---- Responses --------------------------------------------------------------
responses <- plan %>%
  inner_join(students %>% select(section, name, attend, last_class), by = "section",
             relationship = "many-to-many") %>%
  group_by(section, class, name) %>%
  mutate(present = runif(1) < attend[1] & class <= last_class[1]) %>%
  ungroup() %>%
  mutate(answered = present & runif(n()) < 0.92,
         letter = if_else(answered, sample(LETTERS[1:4], n(), replace = TRUE, prob = c(0.45, 0.25, 0.2, 0.1)), NA_character_))

meeting <- responses %>%
  group_by(section, class, name) %>%
  summarise(nq = n(), performance = sum(answered), .groups = "drop") %>%
  mutate(participation = as.integer(performance >= 0.75 * nq),
         total = performance + participation)

# ---- Class-meeting workbooks ------------------------------------------------
pct1 <- function(x) sprintf("%.1f%%", x)   # Question sheets: 100.0%
pct2 <- function(x) sprintf("%.2f%%", x)   # Summary sheet: 92.34%
for (s in sections) {
  out <- file.path(ic_dir, "class_meetings", paste0("sec_", s))
  dir.create(out, recursive = TRUE, showWarnings = FALSE)
  for (n in classes) {
    hm <- strsplit(start_time[[s]], ":")[[1]]
    stamp <- fmt_ic(as.POSIXct(format(class_date(n)), tz = "America/New_York") +
                    as.numeric(hm[1]) * 3600 + (as.numeric(hm[2]) + sample(0:4, 1)) * 60)
    p <- filter(plan, section == s, class == n)
    r <- filter(responses, section == s, class == n) %>% arrange(name, q)
    m <- filter(meeting, section == s, class == n) %>% arrange(name)
    nq <- nrow(p); ns <- nrow(m)
    wb <- createWorkbook()

    # Summary
    addWorksheet(wb, "Summary")
    writeData(wb, "Summary", paste0("Summary - Class ", n, " - Poll"), startRow = 1)
    writeData(wb, "Summary", stamp, startRow = 2)
    writeData(wb, "Summary", "Activity Scores", startRow = 5, startCol = 2)
    writeData(wb, "Summary", "Question Results", startRow = 5, startCol = 5)
    writeData(wb, "Summary", t(c("Name", "Total", "Performance", "Participation",
                                 paste0("Question ", seq_len(nq), "\nMultiple Choice"))),
              startRow = 6, colNames = FALSE)
    pts <- matrix(as.integer(r$answered), nrow = ns, byrow = TRUE)
    writeData(wb, "Summary", data.frame(m$name, m$total, m$performance, m$participation),
              startRow = 8, colNames = FALSE)
    for (k in seq_len(nq)) for (i in seq_len(ns)) {
      writeData(wb, "Summary", if (pts[i, k] == 1) 1 else "NR", startRow = 7 + i, startCol = 4 + k)
    }
    # Averages over students who answered something; the Total % is the mean of the
    # Performance % and Participation % (real example: 92.67% and 92.00% -> 92.34%)
    q_avg <- colSums(pts) / pmax(colSums(pts), 1)          # responders all earn the point
    here <- m$performance > 0
    avg <- function(x) if (any(here)) mean(x[here]) else 0
    perf_pct <- 100 * avg(m$performance) / nq
    part_pct <- 100 * avg(m$participation)
    writeData(wb, "Summary", t(c(nq + 1, nq, 1, rep(1, nq))), startRow = 8 + ns, startCol = 2, colNames = FALSE)
    writeData(wb, "Summary", t(round(c(avg(m$total), avg(m$performance), avg(m$participation), q_avg), 2)),
              startRow = 9 + ns, startCol = 2, colNames = FALSE)
    writeData(wb, "Summary", t(pct2(c((perf_pct + part_pct) / 2, perf_pct, part_pct, 100 * q_avg))),
              startRow = 10 + ns, startCol = 2, colNames = FALSE)
    writeData(wb, "Summary", c("Possible Points", "Average Points", "Average %"), startRow = 8 + ns)

    # Question sheets
    for (k in seq_len(nq)) {
      sh <- paste("Question", k)
      rk <- filter(r, q == k)
      n_resp <- sum(rk$answered)
      addWorksheet(wb, sh)
      writeData(wb, sh, c(paste0("Question ", k, " - Class ", n, " - Poll"), stamp, "Question Type: Multiple Choice"), startRow = 1)
      writeData(wb, sh, data.frame(c("Correct Answers", NA, "Responses", "No Responses", NA, "Possible Points", "Average Points", "Average %"),
                                   c("N/A", NA, n_resp, ns - n_resp, NA, 1, if (n_resp > 0) 1 else 0, pct1(if (n_resp > 0) 100 else 0))),
                startRow = 5, colNames = FALSE)
      writeData(wb, sh, data.frame(c("Name", rk$name), c("Response", if_else(is.na(rk$letter), "NR", rk$letter))),
                startRow = 14, colNames = FALSE)
      # a re-vote shows the very same slide as the first vote
      shown <- if (is.na(p$revote_of[k])) k else p$revote_of[k]
      img <- file.path(img_dir, sprintf("s%s_c%d_q%d.jpg", s, n, shown))
      if (!file.exists(img)) make_slide(img, s, shown, bank$text[p$qid[k]], bank$choices[[p$qid[k]]])
      insertImage(wb, sh, img, startRow = 1, startCol = 4, width = 6.4, height = 3.6)
    }

    suffix <- if (s == "5" && n == 7) " (1)" else ""   # a repeated download
    saveWorkbook(wb, file.path(out, paste0("iClicker_Class", n, "-Poll", suffix, ".xlsx")), overwrite = TRUE)
  }
}

# ---- Gradebooks -------------------------------------------------------------
possible <- plan %>% distinct(section, class, q) %>% count(section, class) %>% mutate(possible = n + 1)
gb_name <- paste0("iClicker_GradesExport_Reef_", format(download, "%m-%d-%y"))
suffixes <- c("2" = "", "5" = " (1)", "11" = " (2)")      # same-day downloads of three sections
for (s in sections) {
  out <- file.path(ic_dir, "gradebook", paste0("sec_", s))
  dir.create(out, recursive = TRUE, showWarnings = FALSE)
  m <- filter(meeting, section == s)
  per_class <- m %>%
    mutate(cell = if_else(performance == 0, "No Response", as.character(total)),
           col = paste0("Class ", class, " - Poll")) %>%
    select(name, col, cell) %>%
    pivot_wider(names_from = col, values_from = cell) %>%
    select(name, all_of(paste0("Class ", rev(classes), " - Poll")))
  tot <- m %>% group_by(name) %>%
    summarise(Total = sum(total), Performance = sum(performance), Participation = sum(participation))
  max_pts <- sum(possible$possible[possible$section == s])
  gb <- students %>%
    filter(section == s) %>%
    left_join(tot, by = "name") %>%
    left_join(per_class, by = "name") %>%
    arrange(name) %>%
    mutate(pct = sprintf("%.2f%%", 100 * Total / max_pts)) %>%
    transmute(Name = name, Email = gb_email, `Student ID` = gb_sid, Total = as.character(Total), `Total %` = pct,
              Performance = as.character(Performance), Participation = as.character(Participation),
              across(starts_with("Class ")))
  write_crlf_csv(names(gb), gb, file.path(out, paste0(gb_name, suffixes[[s]], ".csv")))
}

message("Wrote ", length(sections), " gradebooks and ", length(sections) * length(classes),
        " class-meeting workbooks to ", ic_dir, " (", nrow(students), " students, ",
        sum(students$email %in% roster$Email), " on the dev roster).")
