# Generate fake 2nd-stage quiz exports for development, so Gradescope code is
# never written or tested against real student data. Writes two Gradescope
# exports and one Canvas Classic Quiz "Student Analysis Report" into
# $MA/dev/<term>/2nd_stage/, the category folder that holds both sources (the
# importer must tell them apart by header). Refuses to run unless MA_MODE=dev.
# Leaves any other file in that folder (e.g. the hand-edited "Quiz 1, 2nd stage.csv") alone.
#
# Emails are taken from the dev (hashed) merged roster, so joins to the roster
# can be tested; one student deliberately doesn't match.
#
# Usage (from the project folder): MA_MODE=dev Rscript make_gradescope_dev_data.R
#   Run merge_rosters.R in dev mode first.
#
# Real Gradescope format this mimics (one file per assignment; the instructor
# names the file, so the file name is the assignment name):
#   First Name, Last Name, Student ID (= PSU email), Email, Status (Graded/Missing),
#   Submission ID, Total Score, Max Points, then per question k:
#   Question k Score, Question k Weight, Question k Graded?, Question k Response,
#   Question k Submitted At (M/D/YYYY H:MM). No question text.
#   Missing students: only names, IDs, Status and Max Points filled.
#   A 0-weight question is the collaboration statement (fake text here, no names).
#   CRLF line ends, also inside multi-line quoted responses; no BOM.
# Canvas format: as the dev Initial Survey export (name, id, sis_id, section,
#   section_id, section_sis_id, submitted, attempt, per question "qid: text" +
#   points column headed by points possible, then n correct, n incorrect, score).
# GUESSES (no example seen yet; fix here if the real files differ):
#   - Gradescope rows sorted by last name; scores are whole numbers
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
fmt_et <- function(t) {
  t <- with_tz(t, "America/New_York")
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

make_gradescope <- function(questions, open, due) {
  questions <- mutate(questions, q = row_number())
  max_points <- sum(questions$weight)
  missing <- sample(n_students, 2)
  done <- students[-missing, ]

  # each submitter starts once, then answers (and revises) questions within ~30 minutes
  start <- open + runif(nrow(done), 0, as.numeric(due - open, units = "secs") - 3600)
  answers <- crossing(i = seq_len(nrow(done)), q = questions$q) %>%
    left_join(questions, by = "q") %>%
    mutate(
      ability = done$ability[i],
      correct = rbinom(n(), 1, ability) == 1,
      score = case_when(type == "free" ~ rbinom(n(), weight, ability),
                        type == "mc" ~ as.integer(correct),
                        TRUE ~ 0L),
      response = pmap_chr(list(type, choices, correct), function(type, ch, ok) {
        if (type == "mc") if (ok) ch[1] else sample(ch[-1], 1) else sample(ch, 1)
      }),
      at = start[i] + runif(n(), 0, 30 * 60),
      email = done$email[i]
    )

  per_q <- answers %>%
    transmute(email, q, Score = as.character(score), Weight = as.character(weight),
              `Graded?` = "TRUE", Response = response, `Submitted At` = fmt_et(at)) %>%
    pivot_wider(names_from = q, values_from = all_of(fields),
                names_glue = "Question {q} {.value}", names_vary = "slowest")
  totals <- answers %>% group_by(email) %>% summarise(total = sum(score))

  students %>%
    left_join(totals, by = "email") %>%
    left_join(per_q, by = "email") %>%
    arrange(last, first) %>%
    transmute(`First Name` = first, `Last Name` = last, `Student ID` = email, Email = email,
              Status = if_else(is.na(total), "Missing", "Graded"),
              `Submission ID` = if_else(is.na(total), NA_character_,
                                        as.character(sample(420000000:429999999, n()))),
              `Total Score` = as.character(total), `Max Points` = as.character(max_points),
              across(starts_with("Question ")))
}

gradescope <- list(
  "Quiz 2, 2nd stage.csv" = make_gradescope(quiz2, et("2026-10-09 00:00"), et("2026-10-13 23:59")),
  "Quiz 3, 2nd stage.csv" = make_gradescope(quiz3, et("2026-10-30 00:00"), et("2026-11-03 23:59"))
)
for (f in names(gradescope)) write_crlf_csv(names(gradescope[[f]]), gradescope[[f]], file.path(out_dir, f))

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

message("Wrote ", length(gradescope), " Gradescope files and 1 Canvas file to ", out_dir, " (",
        n_students, " students, ", sum(students$email %in% roster$Email), " on the dev roster).")
