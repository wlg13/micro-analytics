# Generate fake 2nd-stage quiz exports for development, so Gradescope code is
# never written or tested against real student data. Writes, for two quizzes,
# Gradescope's three exports, plus one Canvas Classic Quiz "Student Analysis
# Report", into $MA/dev/<term>/2nd_stage/, the category folder that holds both
# sources (the importer must tell them apart by content). Refuses to run unless
# MA_MODE=dev. Leaves any other file in that folder alone.
#
# Emails are taken from the dev (hashed) merged roster, so joins to the roster
# can be tested; one student deliberately doesn't match.
#
# Usage (from the project folder): MA_MODE=dev Rscript make_gradescope_dev_data.R
#   Run merge_rosters.R in dev mode first.
#
# Real Gradescope formats this mimics (Review Grades page; files kept as downloaded,
# a browser "(n)" may follow; the assignment name comes from the _scores/_evaluations
# names, underscores for spaces, comma lost; the three link by Submission ID):
#   Download Grades, "<Assignment>_scores.csv": First Name, Last Name, SID (= PSU email),
#     Email, Total Score, Max Points, Status (Graded/Missing), Submission ID, Submission
#     Time (M/D/YYYY H:MM), Lateness (H:M:S), View Count, Submission Count, then one
#     score column per question "k (P.0 pts)".
#   Export Evaluations, "<Assignment>_evaluations.zip": a folder "Econ_104_Fa26_<Assignment>"
#     holding 1.csv, 2.csv ... one per question:
#     Assignment Submission ID, Question Submission ID, First Name, Last Name, SID, Email,
#     Score, Submission Time, one TRUE/FALSE column per rubric item headed by its text,
#     Adjustment, Comments, Grader, Tags.
#   Export Submissions, "submissions.zip" (no assignment name): submission_metadata.csv:
#     First Name, Last Name, Student ID, Email, Status, Submission ID, Total Score,
#     Max Points, then per question k: Question k Score, Weight (= points possible),
#     Graded?, Response, Submitted At; the header ends with a trailing comma.
#   No question text anywhere. A 0-point question is the collaboration statement (fake
#   text here, no names). Per-question times read 3 hours earlier than Download Grades'
#   Submission Time (time zones unknown). CRLF line ends, also inside multi-line
#   quoted responses; no BOM.
# Canvas format: as the dev Initial Survey export (name, id, sis_id, section,
#   section_id, section_sis_id, submitted, attempt, per question "qid: text" +
#   points column headed by points possible, then n correct, n incorrect, score).
# GUESSES (no example seen yet; fix here if the real files differ):
#   - Gradescope rows sorted by last name; scores whole numbers; adjustments "-1"
#   - Missing students: in Download Grades and submission_metadata with blank
#     IDs, times and scores; absent from the evaluation CSVs
#   - One rubric item per student per question
#   - Data rows of submission_metadata also end with a trailing comma
#   - A blank answer: empty Response and Submitted At, rubric "blank / no answer"
#   - Canvas name = "First Last", id = 7-digit number, only submitters listed

library(tidyverse)

source("ma_paths.R")
dirs <- ma_dirs()
if (dirs$mode != "dev") stop("make_gradescope_dev_data.R only writes fake data in dev mode; set MA_MODE=dev.")

out_dir <- file.path(dirs$inputs, "2nd_stage")
if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)

roster_path <- file.path(dirs$outputs, "merged_roster.csv")
if (!file.exists(roster_path)) stop("Run merge_rosters.R in dev mode first: ", roster_path, " not found.")
roster <- read_csv(roster_path, col_types = cols(.default = "c"))

set.seed(104)
n_students <- 20

fake_hash <- function(n) replicate(n, paste(sample(c(0:9, letters[1:6]), 16, replace = TRUE), collapse = ""))

# Eastern M/D/YYYY H:MM, 24-hour, no leading zeros (e.g. 9/15/2026 9:19)
fmt_et <- function(t, tz = "America/New_York") {
  t <- with_tz(t, tz)
  paste0(month(t), "/", mday(t), "/", year(t), " ", hour(t), ":", sprintf("%02d", minute(t)))
}
et <- function(x) as.POSIXct(x, tz = "America/New_York")

# One CSV line per row, quoting only fields that need it; NA written as empty.
# CRLF line ends, as in the real exports (responses carry their own CRLFs).
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

first_names <- c("Alex", "Jordan", "Taylor", "Morgan", "Casey", "Riley", "Avery", "Quinn",
                 "Jamie", "Drew", "Reese", "Skyler", "Parker", "Rowan", "Emerson", "Hayden")
last_names <- c("Testperson", "Sample", "Example", "Placeholder", "Fakename", "Mockley",
                "Dummy", "Synthetic", "Madeup", "Fictional", "Notreal", "Specimen")

# 19 students on the dev roster, 1 not (e.g. added after the roster download)
students <- roster %>%
  filter(!is.na(Email)) %>%
  slice_sample(n = n_students - 1) %>%
  transmute(email = Email, section = Section)
students <- bind_rows(students, tibble(email = fake_hash(1), section = "2")) %>%
  mutate(first = sample(first_names, n_students, replace = TRUE),
         last = sample(last_names, n_students, replace = TRUE),
         canvas_id = as.character(sample(7000000:7999999, n_students)),
         ability = runif(n_students, 0.35, 1))

# ---- Question pools ---------------------------------------------------------
# type: free (graded explanation), collab (0-weight collaboration statement),
# mc (choice text; the first choice is correct)
collab_text <- c("I worked alone.", "I worked by myself using my notes from class.",
                 "I used only the textbook and my notes.", "I did not work with anyone.")
free <- function(weight, ...) tibble(type = "free", weight = weight, choices = list(c(...)))
mc <- function(...) tibble(type = "mc", weight = 1, choices = list(c(...)))
collab <- tibble(type = "collab", weight = 0, choices = list(collab_text))

quiz2 <- bind_rows(
  free(5, "B. The CPI overstates inflation\r\nWhen prices rise, people substitute toward cheaper goods, but the CPI uses a fixed basket, so it overstates the rise in the cost of living.",
          "B) The CPI overstates inflation, because it does not fully account for new goods and quality improvements.",
          "A. The CPI understates inflation\r\nPrices of things I buy, like rent and food, rose more than the CPI.",
          "The CPI overstates inflation because of substitution bias, quality bias and new goods bias. "),
  free(5, "C. 5%\r\nThe unemployment rate is the number unemployed divided by the labor force, 5 / 100 = 5%.",
          "C) 5% because 5 million unemployed out of a labor force of 100 million.",
          "B. 4%\r\nI divided the unemployed by the adult population.",
          "5 percent, since the unemployment rate = unemployed / labor force."),
  collab,
  mc("frictional unemployment", "cyclical unemployment", "structural unemployment"),
  mc("real GDP grew", "real GDP fell", "real GDP did not change"),
  mc("it would rise", "it would fall", "it would stay the same"),
  mc("a discouraged worker", "an unemployed worker", "an employed worker"),
  mc("the labor force participation rate fell", "the unemployment rate rose", "nothing changed")
)

quiz3 <- bind_rows(
  free(5, "A. Higher productivity\r\nIn the long run, living standards depend on output per worker, which rises with physical capital, human capital and technology.",
          "A) productivity, because more output per hour lets each person consume more.",
          "C. More workers\r\nMore people working produces more goods.",
          "Productivity growth is the key, driven by technology and investment in capital. "),
  free(5, "B. A recession\r\nReal GDP fell for two quarters and unemployment rose.",
          "B) recession, since output fell and unemployment increased.",
          "D. An expansion\r\nPrices were still rising.",
          "This is a recession because real GDP declined."),
  free(3, "Aggregate demand shifts left, so output and the price level fall.",
          "AD shifts to the left; real GDP falls and prices fall in the short run.",
          "Aggregate supply shifts left.",
          "AD decreases\r\nOutput falls below potential, and unemployment rises."),
  collab,
  mc("the rule of 70", "the rule of 72", "the growth accounting rule"),
  mc("about 35 years", "about 70 years", "about 2 years"),
  mc("a peak", "a trough", "an expansion"),
  mc("unemployment rises", "unemployment falls", "unemployment does not change"),
  mc("aggregate demand shifts right", "aggregate supply shifts left", "potential GDP falls"),
  mc("the price level rises", "the price level falls", "the price level stays the same"),
  mc("interest rates", "tariffs", "the minimum wage")
)

# ---- Gradescope -------------------------------------------------------------
fields <- c("Score", "Weight", "Graded?", "Response", "Submitted At")
grader <- "Pat Instructor"
pacific <- "America/Los_Angeles"

# Rubric items for a question, best first, with their points. Free-response:
# the correct answer, the right letter without a good explanation, anything else,
# blank. Multiple choice: the correct choice, anything else, blank.
rubric <- function(type, weight, choices) {
  switch(type,
         free = tibble(item = c(sub("\r\n.*", "", choices[1]),
                                paste0('"', substr(choices[1], 1, 1), '" with no explanation or an incorrect explanation'),
                                "Any other answer.", "blank / no answer"),
                       points = c(weight, weight - 2, 0, 0)),
         mc = tibble(item = c(choices[1], "Any other answer.", "blank / no answer"), points = c(1, 0, 0)),
         collab = tibble(item = c("Collaboration statement given", "blank / no answer"), points = c(0, 0)))
}

lateness <- function(sec) sprintf("%d:%02d:%02d", sec %/% 3600, sec %/% 60 %% 60, sec %% 60)

make_gradescope <- function(questions, title, open, due) {
  questions <- mutate(questions, q = row_number(), rubric = pmap(list(type, weight, choices), rubric))
  max_points <- sum(questions$weight)
  missing <- sample(n_students, 2)
  done <- students[-missing, ] %>%
    mutate(sub_id = as.character(sample(420000000:429999999, n())),
           views = sample(0:3, n(), replace = TRUE, prob = c(0.5, 0.3, 0.15, 0.05)),
           n_sub = sample(1:2, n(), replace = TRUE, prob = c(0.85, 0.15)))

  # each submitter starts once, then answers (and revises) questions within ~30 minutes;
  # one starts after the deadline (late)
  start <- open + runif(nrow(done), 0, as.numeric(due - open, units = "secs") - 3600)
  start[1] <- due + runif(1, 60, 7200)
  answers <- crossing(i = seq_len(nrow(done)), q = questions$q) %>%
    left_join(questions, by = "q") %>%
    mutate(
      ability = done$ability[i],
      blank = runif(n()) < 0.03,
      # rubric item received: 1 = best; free-response can be half right
      item = case_when(blank ~ map_int(rubric, nrow),
                       type == "collab" ~ 1L,
                       runif(n()) < ability ~ 1L,
                       type == "free" & runif(n()) < 0.5 ~ 2L,
                       TRUE ~ map_int(rubric, nrow) - 1L),
      adjustment = if_else(type == "free" & !blank & runif(n()) < 0.05, -1, NA_real_),
      score = map2_dbl(rubric, item, ~ .x$points[.y]) + coalesce(adjustment, 0),
      response = pmap_chr(list(type, choices, item, blank), function(type, ch, it, b) {
        if (b) NA_character_
        else if (type == "free") ch[c(1, 4, 3)][it]          # correct, short, wrong
        else if (type == "mc") if (it == 1) ch[1] else sample(ch[-1], 1)
        else sample(ch, 1)
      }),
      at = start[i] + runif(n(), 0, 30 * 60),
      qsub_id = as.character(sample(4110000000:4129999999, n())),
      email = done$email[i]
    )
  last_at <- answers %>% group_by(i) %>% summarise(at = max(at[!blank]))   # blank answers carry no time
  done <- mutate(done, final = last_at$at, late = pmax(0, round(as.numeric(final - due, units = "secs"))),
                 total = tapply(answers$score, answers$i, sum))
  everyone <- students %>% left_join(select(done, email, sub_id, views, n_sub, final, late, total), by = "email") %>%
    arrange(last, first)
  num <- function(x) if_else(is.na(x), NA_character_, as.character(x))

  # Export Submissions: zip holding submission_metadata.csv. Per-question times are
  # 3 hours behind Download Grades in the real exports, so they're written as Pacific
  # clock time here (a guess); the header and every row end with a trailing comma.
  per_q <- answers %>%
    transmute(email, q, Score = num(score), Weight = num(weight), `Graded?` = "TRUE",
              Response = response, `Submitted At` = if_else(blank, NA_character_, fmt_et(at, pacific))) %>%
    pivot_wider(names_from = q, values_from = all_of(fields),
                names_glue = "Question {q} {.value}", names_vary = "slowest")
  submissions <- everyone %>%
    left_join(per_q, by = "email") %>%
    transmute(`First Name` = first, `Last Name` = last, `Student ID` = email, Email = email,
              Status = if_else(is.na(total), "Missing", "Graded"), `Submission ID` = sub_id,
              `Total Score` = num(total), `Max Points` = as.character(max_points),
              across(starts_with("Question ")), trailing = NA_character_)

  # Download Grades: one CSV, per-question score columns "k (P.0 pts)"
  scores <- answers %>%
    transmute(email, col = sprintf("%d (%.1f pts)", q, weight), score = num(score)) %>%
    pivot_wider(names_from = col, values_from = score)
  grades <- everyone %>%
    left_join(scores, by = "email") %>%
    transmute(`First Name` = first, `Last Name` = last, SID = email, Email = email,
              `Total Score` = num(total), `Max Points` = as.character(max_points),
              Status = if_else(is.na(total), "Missing", "Graded"), `Submission ID` = sub_id,
              `Submission Time` = if_else(is.na(final), NA_character_, fmt_et(final)),
              `Lateness (H:M:S)` = if_else(is.na(late), NA_character_, lateness(late)),
              `View Count` = num(views), `Submission Count` = num(n_sub),
              across(matches("^[0-9]+ \\(")))

  # Export Evaluations: zip of 1.csv, 2.csv ... (submitters only), one TRUE/FALSE
  # column per rubric item headed by the item's text
  evaluations <- map(questions$q, function(k) {
    items <- questions$rubric[[k]]$item
    a <- answers %>% filter(q == k) %>% left_join(select(done, email, sub_id, first, last), by = "email") %>%
      arrange(last, first)
    marks <- map(seq_along(items), ~ if_else(a$item == .x, "TRUE", "FALSE")) %>% set_names(items) %>% as_tibble()
    bind_cols(
      tibble(`Assignment Submission ID` = a$sub_id, `Question Submission ID` = a$qsub_id,
             `First Name` = a$first, `Last Name` = a$last, SID = a$email, Email = a$email,
             Score = num(a$score), `Submission Time` = if_else(a$blank, NA_character_, fmt_et(a$at, pacific))),
      marks,
      tibble(Adjustment = num(a$adjustment),
             Comments = if_else(is.na(a$adjustment), NA_character_, "Explanation incomplete."),
             Grader = grader, Tags = NA_character_))
  })
  list(title = title, submissions = submissions, grades = grades, evaluations = evaluations)
}

# File names as Gradescope gives them; (n) = a repeated browser download
write_gradescope <- function(g, sfx, submissions_zip) {
  stem <- gsub("[ ,]+", "_", g$title)                       # "Quiz 2, 2nd Stage" -> Quiz_2_2nd_Stage
  write_crlf_csv(names(g$grades), g$grades, file.path(out_dir, paste0(stem, "_scores", sfx, ".csv")))

  # the question CSVs sit in a folder named for the course and assignment
  tmp <- tempfile(); folder <- paste0("Econ_104_Fa26_", stem); dir.create(file.path(tmp, folder), recursive = TRUE)
  for (k in seq_along(g$evaluations)) write_crlf_csv(names(g$evaluations[[k]]), g$evaluations[[k]], file.path(tmp, folder, paste0(k, ".csv")))
  zip::zip(file.path(out_dir, paste0(stem, "_evaluations", sfx, ".zip")), folder, root = tmp)

  tmp <- tempfile(); dir.create(tmp)
  header <- names(g$submissions); header[length(header)] <- ""
  write_crlf_csv(header, g$submissions, file.path(tmp, "submission_metadata.csv"))
  zip::zip(file.path(out_dir, submissions_zip), "submission_metadata.csv", root = tmp)
}

unlink(file.path(out_dir, c("Quiz 2, 2nd stage.csv", "Quiz 3, 2nd stage.csv")))   # earlier layout
gradescope <- list(make_gradescope(quiz2, "Quiz 2, 2nd Stage", et("2026-10-09 00:00"), et("2026-10-13 23:59")),
                   make_gradescope(quiz3, "Quiz 3, 2nd Stage", et("2026-10-30 00:00"), et("2026-11-03 23:59")))
write_gradescope(gradescope[[1]], "", "submissions.zip")
write_gradescope(gradescope[[2]], "(1)", "submissions(1).zip")

# ---- Canvas Classic Quiz (same folder, so source detection can be tested) ---
sections <- tribble(
  ~section, ~canvas_section,                                                                         ~section_id, ~section_sis_id,
  "2",      "ECON 104, Section 002: Macroeconomic Analysis (22681--UP---P-ECON----104-------002-)", "2664009",   "s22681--16246",
  "5",      "ECON 104, Section 005: Macroeconomic Analysis (22681--UP---P-ECON----104-------005-)", "2664075",   "s22681--16272",
  "11",     "ECON 104, Section 011: Macroeconomic Analysis (22681--UP---P-ECON----104-------011-)", "2672183",   "s22681--29328"
)
canvas_qs <- tibble(
  qid = as.character(101552301:101552305),
  text = c("If aggregate demand shifts left, in the short run real GDP ___.",
           "Which of the following shifts short-run aggregate supply?",
           "A fall in stock prices, holding all else constant, shifts aggregate demand ___.",
           "The economy is at a point above potential GDP. Which is most likely?",
           "Which of the following is most accurate?"),
  choices = list(c("falls", "rises", "does not change"),
                 c("a change in oil prices", "a change in the price level", "a change in government purchases"),
                 c("to the left", "to the right", "neither way"),
                 c("inflation will rise", "unemployment is high", "real GDP is falling"),
                 c("output and unemployment move in opposite directions", "output and unemployment move together",
                   "output and unemployment are unrelated"))
)
takers <- students[sample(n_students, n_students - 3), ] %>% left_join(sections, by = "section")
submitted <- et("2026-11-02 00:00") + runif(nrow(takers), 0, 4 * 86400)
canvas_rows <- tibble(name = paste(takers$first, takers$last), id = takers$canvas_id, sis_id = takers$email,
                      section = takers$canvas_section, section_id = takers$section_id,
                      section_sis_id = takers$section_sis_id,
                      submitted = format(with_tz(submitted, "UTC"), "%Y-%m-%d %H:%M:%S UTC"), attempt = "1")
n_right <- integer(nrow(takers))
for (k in seq_len(nrow(canvas_qs))) {
  ch <- canvas_qs$choices[[k]]
  ok <- rbinom(nrow(takers), 1, takers$ability) == 1
  n_right <- n_right + ok
  canvas_rows[[paste0("a", k)]] <- if_else(ok, ch[1], map_chr(ok, ~ sample(ch[-1], 1)))
  canvas_rows[[paste0("p", k)]] <- as.character(as.integer(ok))
}
canvas_rows <- mutate(canvas_rows, `n correct` = as.character(n_right),
                      `n incorrect` = as.character(nrow(canvas_qs) - n_right), score = `n correct`)
canvas_header <- c("name", "id", "sis_id", "section", "section_id", "section_sis_id", "submitted", "attempt",
                   rbind(paste0(canvas_qs$qid, ": ", canvas_qs$text), "1"),
                   "n correct", "n incorrect", "score")
write_crlf_csv(canvas_header, canvas_rows, file.path(out_dir, "Quiz 4, 2nd stage Quiz Student Analysis Report.csv"))

message("Wrote ", length(gradescope), " Gradescope assignments (3 exports each) and 1 Canvas file to ", out_dir, " (",
        n_students, " students, ", sum(students$email %in% roster$Email), " on the dev roster).")
