# Generate fake roster inputs for development, so code is never written or
# tested against real student data. Mirrors the structure of the real exports
# that merge_rosters.R reads from $MA/inputs.
#
# Usage: Rscript make_dev_data.R [output dir]   (default: dev_data)
# Then develop with MA pointing at that dir, e.g. MA=dev_data Rscript merge_rosters.R

library(tidyverse)
library(writexl)

args <- commandArgs(trailingOnly = TRUE)
dev_dir <- if (length(args) > 0) args[1] else "dev_data"
inputs_dir <- file.path(dev_dir, "inputs")
if (!dir.exists(inputs_dir)) dir.create(inputs_dir, recursive = TRUE)

set.seed(104)
n <- 60

first_names <- c("Alex", "Jordan", "Taylor", "Morgan", "Casey", "Riley", "Avery", "Quinn",
                 "Jamie", "Drew", "Reese", "Skyler", "Parker", "Rowan", "Emerson", "Hayden")
last_names <- c("Testperson", "Sample", "Example", "Placeholder", "Fakename", "Mockley",
                "Dummy", "Synthetic", "Madeup", "Fictional", "Notreal", "Specimen")

students <- tibble(
  ID = sprintf("0%08d", sample(1e6:9e6, n)),  # leading 0: clearly not a real PSU ID
  first = sample(first_names, n, replace = TRUE),
  last = sample(last_names, n, replace = TRUE),
  campus_id = sprintf("zz%04d", sample(1000:9999, n)),
  email = paste0(campus_id, "@example.edu"),
  section = sample(c(2, 5, 11), n, replace = TRUE),
  status = sample(c(NA, "Late Dropped", "Withdrawn"), n, replace = TRUE, prob = c(0.9, 0.05, 0.05))
)

lionpath <- students %>%
  transmute(
    Notify = NA_character_,
    ID,
    Name = paste0(last, ",", first),
    Pronouns = sample(c(NA, "She/Her", "He/Him", "They/Them"), n, replace = TRUE),
    Units = 3,
    `Program and Plan` = sample(c("Economics (BS)", "Business Undeclared", "Division of Undergraduate Studies",
                                  "Engineering Undeclared", "Political Science (BA)"), n, replace = TRUE),
    Level = sample(c("Freshman", "Sophomore", "Junior", "Senior"), n, replace = TRUE),
    `Status Note` = status,
    Section = section
  )

# Activity Insights: a few students missing (tests unmatched rows) and a few
# exact-duplicate rows (tests distinct()).
activity_insights <- students %>%
  slice_sample(n = n - 3) %>%
  transmute(`PSU ID` = ID, `Campus ID` = campus_id, Email = email,
            `Last Name` = last, `First Name` = first)
activity_insights <- bind_rows(activity_insights, slice_sample(activity_insights, n = 4))

# Canvas survey report: sis_id is the email; some upper-cased (tests the
# case-insensitive join) and some students never submitted.
q_col <- "100874763: Are you the first person in your family to attend college?"
survey <- students %>%
  slice_sample(prop = 0.8) %>%
  transmute(
    name = paste(first, last),
    id = sample(100000:999999, n()),
    sis_id = if_else(runif(n()) < 0.2, toupper(email), email),
    section = paste("ECON 104 Section", section),
    submitted = "2026-08-30 12:00:00 UTC",
    !!q_col := sample(c("Yes", "No", "Not sure"), n(), replace = TRUE, prob = c(0.3, 0.6, 0.1)),
    score = 1
  )

write_xlsx(lionpath, file.path(inputs_dir, "roster_LionPATH.xlsx"))
write_xlsx(activity_insights, file.path(inputs_dir, "roster_Activity_Insights.xlsx"))
write_csv(survey, file.path(inputs_dir, "Initial Survey Survey Student Analysis Report.csv"), na = "")

message("Wrote fake inputs to ", normalizePath(inputs_dir))
