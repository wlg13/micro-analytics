# Input and output folders under $MA, chosen by $MA_MODE and $MA_TERM:
#   MA_MODE unset (or "production"): $MA/<term>/     -> $MA/<term>/outputs/      (real data)
#   MA_MODE=dev:                     $MA/dev/<term>/ -> $MA/dev/<term>/outputs/  (hashed data)
# <term> is a term folder such as Fa26 or Sp27: $MA_TERM if set, otherwise the
# only term folder present (it stops and lists them if there are several).
# Every script should get its folders from ma_dirs() so the switch applies everywhere.

TERM_PATTERN <- "^(Fa|Sp|Su)[0-9]{2}$"

ma_dirs <- function() {
  ma_dir <- Sys.getenv("MA")
  if (ma_dir == "") stop("Environment variable MA is not set (should point to the data directory).")

  mode <- tolower(trimws(Sys.getenv("MA_MODE")))
  if (mode %in% c("", "production")) {
    mode <- "production"
    base <- ma_dir
  } else if (mode == "dev") {
    base <- file.path(ma_dir, "dev")
  } else {
    stop('MA_MODE must be "dev" or unset (production), not "', Sys.getenv("MA_MODE"), '".')
  }

  term <- trimws(Sys.getenv("MA_TERM"))
  if (term == "") {
    terms <- grep(TERM_PATTERN, list.dirs(base, full.names = FALSE, recursive = FALSE), value = TRUE)
    if (length(terms) == 1) {
      term <- terms
    } else if (length(terms) == 0) {
      stop("No term folder (e.g. Fa26) found in ", base, ".")
    } else {
      stop("Several term folders in ", base, " (", paste(terms, collapse = ", "),
           "). Set MA_TERM to pick one, e.g. Sys.setenv(MA_TERM = \"", tail(sort(terms), 1), "\").")
    }
  } else if (!grepl(TERM_PATTERN, term)) {
    stop('MA_TERM must look like Fa26, Sp27 or Su27, not "', term, '".')
  }

  dirs <- list(
    mode = mode,
    term = term,
    inputs = file.path(base, term),
    outputs = file.path(base, term, "outputs")
  )
  if (!dir.exists(dirs$inputs)) stop("Input folder not found: ", dirs$inputs)
  if (!dir.exists(dirs$outputs)) dir.create(dirs$outputs, recursive = TRUE)

  shown <- if (mode == "dev") file.path("dev", term) else term
  message("Mode: ", mode, ", term ", term, " (", shown, " -> ", file.path(shown, "outputs"), ")")
  dirs
}
