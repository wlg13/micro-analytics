# Merge LionPATH_roster.xlsx with roster_details.xlsx by ID / PSU ID

library(tidyverse)
library(readxl)

ma_dir <- Sys.getenv("MA")
if (ma_dir == "") stop("Environment variable MA is not set (should point to the data directory).")

inputs_dir <- file.path(ma_dir, "inputs")
outputs_dir <- file.path(ma_dir, "outputs")

roster_details_path <- file.path(inputs_dir, "roster_details.xlsx")
lionpath_path <- file.path(inputs_dir, "LionPATH_roster.xlsx")

roster_details <- read_excel(roster_details_path)
lionpath <- read_excel(lionpath_path)

message("roster_details.xlsx: ", nrow(roster_details), " rows")
message("LionPATH_roster.xlsx: ", nrow(lionpath), " rows")

merged_roster <- left_join(lionpath, roster_details, by = c("ID" = "PSU ID"))

n_unmatched <- sum(is.na(merged_roster$Email))
message("Merged roster: ", nrow(merged_roster), " rows (", n_unmatched, " with no match in roster_details.xlsx)")

if (!dir.exists(outputs_dir)) dir.create(outputs_dir, recursive = TRUE)

output_path <- file.path(outputs_dir, "merged_roster.csv")
write_csv(merged_roster, output_path)
message("Wrote ", output_path)
