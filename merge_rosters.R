# Merge roster_LionPATH.xlsx with roster_Activity_Insights.xlsx by ID / PSU ID

library(tidyverse)
library(readxl)

ma_dir <- Sys.getenv("MA")
if (ma_dir == "") stop("Environment variable MA is not set (should point to the data directory).")

inputs_dir <- file.path(ma_dir, "inputs")
outputs_dir <- file.path(ma_dir, "outputs")

roster_details_path <- file.path(inputs_dir, "roster_Activity_Insights.xlsx")
lionpath_path <- file.path(inputs_dir, "roster_LionPATH.xlsx")

roster_details <- read_excel(roster_details_path) %>%
  mutate(`PSU ID` = as.character(`PSU ID`)) %>%
  distinct()
lionpath <- read_excel(lionpath_path) %>%
  mutate(ID = as.character(ID)) %>%
  filter(!`Status Note` %in% c("Late Dropped", "Withdrawn"))

message("roster_Activity_Insights.xlsx: ", nrow(roster_details), " rows")
message("roster_LionPATH.xlsx: ", nrow(lionpath), " rows")

merged_roster <- left_join(lionpath, roster_details, by = c("ID" = "PSU ID"))

n_unmatched <- sum(is.na(merged_roster$Email))
message("Merged roster: ", nrow(merged_roster), " rows (", n_unmatched, " with no match in roster_Activity_Insights.xlsx)")

survey_path <- file.path(inputs_dir, "Initial Survey Survey Student Analysis Report.csv")
survey <- read_csv(survey_path, col_select = c(sis_id, starts_with("100874763")), col_types = cols(.default = "c")) %>%
  rename(`First Generation` = starts_with("100874763")) %>%
  mutate(email_key = tolower(sis_id)) %>%
  select(-sis_id)

merged_roster <- merged_roster %>%
  mutate(email_key = tolower(Email)) %>%
  left_join(survey, by = "email_key") %>%
  select(-email_key)

n_no_survey <- sum(!tolower(merged_roster$Email) %in% survey$email_key)
message("Survey: ", nrow(survey), " rows (", n_no_survey, " roster students with no survey submission)")

merged_roster <- merged_roster %>%
  select(`First Name`, `Last Name`, ID, Email, Level, FirstGen = `First Generation`, `Program and Plan`)

if (!dir.exists(outputs_dir)) dir.create(outputs_dir, recursive = TRUE)

output_path <- file.path(outputs_dir, "merged_roster.csv")
write_csv(merged_roster, output_path)
message("Wrote ", output_path)
