# Merge LionPATH_roster.xlsx with roster_details.xlsx by ID / PSU ID

library(tidyverse)
library(readxl)

roster_details_path <- path.expand("~/roster_details.xlsx")
lionpath_path <- path.expand("~/LionPATH_roster.xlsx")

roster_details <- read_excel(roster_details_path)
lionpath <- read_excel(lionpath_path)

message("roster_details.xlsx: ", nrow(roster_details), " rows")
message("LionPATH_roster.xlsx: ", nrow(lionpath), " rows")

merged_roster <- left_join(lionpath, roster_details, by = c("ID" = "PSU ID"))

n_unmatched <- sum(is.na(merged_roster$Email))
message("Merged roster: ", nrow(merged_roster), " rows (", n_unmatched, " with no match in roster_details.xlsx)")

write_csv(merged_roster, "merged_roster.csv")
message("Wrote merged_roster.csv")
