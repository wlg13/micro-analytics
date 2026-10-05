# Input and output folders under $MA, chosen by $MA_MODE:
#   MA_MODE unset (or "production"): inputs/     -> outputs/      (real data)
#   MA_MODE=dev:                     inputs-dev/ -> outputs-dev/  (hashed data)
# Every script should get its folders from ma_dirs() so the switch applies everywhere.

ma_dirs <- function() {
  ma_dir <- Sys.getenv("MA")
  if (ma_dir == "") stop("Environment variable MA is not set (should point to the data directory).")

  mode <- tolower(trimws(Sys.getenv("MA_MODE")))
  if (mode %in% c("", "production")) {
    mode <- "production"
    suffix <- ""
  } else if (mode == "dev") {
    suffix <- "-dev"
  } else {
    stop('MA_MODE must be "dev" or unset (production), not "', Sys.getenv("MA_MODE"), '".')
  }

  dirs <- list(
    mode = mode,
    inputs = file.path(ma_dir, paste0("inputs", suffix)),
    outputs = file.path(ma_dir, paste0("outputs", suffix))
  )
  if (!dir.exists(dirs$inputs)) stop("Input folder not found: ", dirs$inputs)
  if (!dir.exists(dirs$outputs)) dir.create(dirs$outputs, recursive = TRUE)

  message("Mode: ", mode, " (", basename(dirs$inputs), " -> ", basename(dirs$outputs), ")")
  dirs
}
