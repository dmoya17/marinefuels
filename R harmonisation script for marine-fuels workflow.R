###############################################################################
# HSFO marine fuels harmonisation workflow
# Reconciles PRELIM HSFO carbon-intensity outputs with REM / Commercial data
# refinery fuel-oil volumes, prices, and export shares.
#
# Scope:
# - no plots
# - deterministic and reproducible
# - repository-ready structure
# - explicit inputs, outputs, validation, and diagnostics
#
# Main outputs:
# 01_prelim_hsfo_asset_level_<timestamp>.csv
# 02_prelim_hsfo_with_locations_<timestamp>.csv
# 03_rem_hsfo_volumes_<timestamp>.csv
# 04_rem_hsfo_prices_shares_<timestamp>.csv
# 05_hsfo_harmonised_all_rows_<timestamp>.csv
# 06_hsfo_harmonised_analysis_ready_<timestamp>.csv
# 07_hsfo_harmonisation_diagnostics_<timestamp>.csv
# 08_session_info_<timestamp>.txt
#
# Notes:
# - this script assumes PRELIM outputs already exist
# - this script does not run OPGEE, PRELIM, or COPTEM
# - this script is limited to HSFO / fuel-oil harmonisation
###############################################################################

rm(list = ls())

# ---------------------------------------------------------------------------
# 0) required packages
# ---------------------------------------------------------------------------
required_pkgs <- c(
  "dplyr",
  "tidyr",
  "stringr",
  "readr",
  "readxl",
  "purrr",
  "tibble",
  "janitor",
  "countrycode"
)

missing_pkgs <- setdiff(required_pkgs, rownames(installed.packages()))
if (length(missing_pkgs) > 0) {
  stop(
    "Install missing packages before running this script: ",
    paste(missing_pkgs, collapse = ", ")
  )
}

invisible(lapply(required_pkgs, library, character.only = TRUE))

# ---------------------------------------------------------------------------
# 1) user configuration: edit file paths here
# ---------------------------------------------------------------------------
cfg <- list(
  prelim_ci_xlsx = "PATH/Marine fuel well to refinery exit CI v5 (PRELIMv20) LPG lower v2 2023.xlsx",
  prelim_locations_xlsx = "PATH/10 Phase III/00 inputs/refinery locations.xlsx",
  rem_prices_csv = "PATH/10 Phase III/00 inputs/all_refineries_product_prices_export.csv",
  rem_files = c(
    "PATH/01 inputs/rem-_russia_and_caspian_ncm_benchmarking_2023/russia_and_caspian_financialmodel_2023.xlsx",
    "PATH/01 inputs/rem-_north_america_ncm_benchmarking_2023/northamerica_financialmodel_2023.xlsx",
    "PATH/01 inputs/rem-_middle_east_ncm_benchmarking_2023/middleeast_financialmodel_2023.xlsx",
    "PATH/01 inputs/rem-_latin_america_ncm_benchmarking_2023/latin-america_financialmodel_2023.xlsx",
    "PATH/01 inputs/rem-_europe_ncm_benchmarking_2023/europe_financialmodel_2023.xlsx",
    "PATH/01 inputs/rem-_china_ncm_benchmarking_2023/china_financialmodel_2023.xlsx",
    "PATH/01 inputs/rem-_asia_pacific_ex-china_ncm_benchmarking_2023/asiapacific_ex_china_financialmodel_2023.xlsx",
    "PATH/01 inputs/rem-_africa_ncm_benchmarking_2023/africa_financialmodel_2023.xlsx"
  ),
  out_dir = "PATH/02 outputs/03 v2 outputs"
)

dir.create(cfg$out_dir, showWarnings = FALSE, recursive = TRUE)

# ---------------------------------------------------------------------------
# 2) helper functions
# ---------------------------------------------------------------------------
log_step <- function(x) message("\n", x)

assert_file_exists <- function(path) {
  if (!file.exists(path)) stop("File not found: ", path)
}

assert_files_exist <- function(paths) {
  invisible(lapply(paths, assert_file_exists))
}

normalize_name <- function(x) {
  x %>%
    as.character() %>%
    tolower() %>%
    stringr::str_replace_all("[[:punct:]]", " ") %>%
    stringr::str_squish()
}

safe_num <- function(x) suppressWarnings(as.numeric(x))

weighted_mean_na <- function(x, w) {
  ok <- !is.na(x) & !is.na(w) & is.finite(x) & is.finite(w) & w > 0
  if (!any(ok)) return(NA_real_)
  sum(x[ok] * w[ok]) / sum(w[ok])
}

mean_positive_na <- function(x) {
  ok <- !is.na(x) & is.finite(x) & x > 0
  if (!any(ok)) return(NA_real_)
  mean(x[ok])
}

first_existing_col <- function(df, candidates, required = TRUE, label = "column") {
  hit <- intersect(candidates, names(df))
  if (length(hit) == 0) {
    if (required) {
      stop(
        "No ", label, " found. Candidates tested: ",
        paste(candidates, collapse = ", "),
        "\nAvailable columns: ",
        paste(names(df), collapse = ", ")
      )
    } else {
      return(NA_character_)
    }
  }
  hit[1]
}

pick_col_by_simplified_name <- function(df, pattern, required = TRUE, label = "column") {
  nm <- names(df)
  nm_simpl <- setNames(stringr::str_replace_all(nm, "[^a-z0-9]", ""), nm)
  hit <- names(nm_simpl)[stringr::str_detect(nm_simpl, regex(pattern, ignore_case = TRUE))]
  if (length(hit) == 0) {
    if (required) {
      stop(
        "Could not detect ", label, " using pattern: ", pattern,
        "\nAvailable columns: ", paste(nm, collapse = ", ")
      )
    } else {
      return(NA_character_)
    }
  }
  hit[1]
}

fill_region_from_country <- function(df, region_col = "region", country_col = "analysis_country") {
  stopifnot(region_col %in% names(df), country_col %in% names(df))
  lkp <- df %>%
    filter(!is.na(.data[[region_col]]), !is.na(.data[[country_col]])) %>%
    distinct(.data[[country_col]], .data[[region_col]])
  names(lkp) <- c(country_col, "region_lkp")
  
  df %>%
    left_join(lkp, by = country_col) %>%
    mutate(
      !!region_col := dplyr::coalesce(.data[[region_col]], .data[["region_lkp"]])
    ) %>%
    select(-region_lkp)
}

# ---------------------------------------------------------------------------
# 3) readers
# ---------------------------------------------------------------------------
read_prelim_hsfo <- function(file) {
  assert_file_exists(file)
  
  log_step("[1] Loading PRELIM CI")
  df <- readxl::read_excel(file, sheet = 1) %>%
    janitor::clean_names()
  
  if ("x1" %in% names(df)) names(df)[names(df) == "x1"] <- "refinery_id"
  if ("mj_bbl" %in% names(df)) names(df)[names(df) == "mj_bbl"] <- "energy_density_lhv_mj_bbl"
  
  drop_cols <- c(
    "x8",
    "x11",
    "x18",
    "hsfo_domestic_use_bbl_d",
    "hsfo_import_bbl_d",
    "fuel_trade_shipping_ci_g_co2e_mj",
    "combustion_ci_g_co2e_mj"
  )
  
  df <- df %>% select(-any_of(drop_cols))
  
  if (!"refinery_id" %in% names(df)) {
    stop("PRELIM CI file must contain refinery_id or x1.")
  }
  
  name_col <- first_existing_col(
    df,
    candidates = c("prelim_refinery_name", "refinery_name", "refinery", "name"),
    required = FALSE,
    label = "PRELIM refinery name"
  )
  
  country_col <- first_existing_col(
    df,
    candidates = c("country", "country_1", "country_name"),
    required = FALSE,
    label = "PRELIM country"
  )
  
  if (is.na(name_col)) {
    df$prelim_refinery_name <- NA_character_
  } else {
    df$prelim_refinery_name <- as.character(df[[name_col]])
  }
  
  if (is.na(country_col)) {
    df$country_prelim <- NA_character_
  } else {
    df$country_prelim <- as.character(df[[country_col]])
  }
  
  needed <- c(
    "hsfo_volume_bbl_d",
    "hsfo_refining_ci_g_co2e_mj",
    "hsfo_total_before_shipping"
  )
  
  missing_needed <- setdiff(needed, names(df))
  if (length(missing_needed) > 0) {
    stop("PRELIM CI file is missing required HSFO columns: ", paste(missing_needed, collapse = ", "))
  }
  
  df %>%
    mutate(
      hsfo_volume_bbl_d = safe_num(hsfo_volume_bbl_d),
      hsfo_refining_ci_g_co2e_mj = safe_num(hsfo_refining_ci_g_co2e_mj),
      hsfo_total_before_shipping = safe_num(hsfo_total_before_shipping)
    ) %>%
    select(-matches("^lpg_")) %>%
    filter(
      !is.na(hsfo_volume_bbl_d),
      !is.na(hsfo_refining_ci_g_co2e_mj),
      !is.na(hsfo_total_before_shipping)
    ) %>%
    rename(prelim_hsfo_volume_bpd = hsfo_volume_bbl_d)
}

read_prelim_locations <- function(file, prelim_hsfo) {
  assert_file_exists(file)
  
  log_step("[2] Loading refinery locations and joining to PRELIM")
  loc <- readxl::read_excel(file, sheet = 1) %>%
    janitor::clean_names()
  
  id_col <- first_existing_col(loc, c("refinery_id"), label = "refinery_id")
  lat_col <- first_existing_col(loc, c("latitude", "lat"), label = "latitude")
  lon_col <- first_existing_col(loc, c("longitude", "lon", "lng"), label = "longitude")
  
  loc <- loc %>%
    rename(
      refinery_id = !!id_col,
      lat_ref = !!lat_col,
      lon_ref = !!lon_col
    )
  
  loc %>%
    left_join(prelim_hsfo, by = "refinery_id") %>%
    filter(!is.na(refinery_id))
}

read_wm_sheet <- function(file) {
  assert_file_exists(file)
  
  df_raw <- readxl::read_excel(file, sheet = "Refinery Output & Prices", skip = 4) %>%
    janitor::clean_names() %>%
    tibble::as_tibble()
  
  col_region <- first_existing_col(df_raw, c("region"), required = FALSE, label = "region")
  col_country <- first_existing_col(df_raw, c("country"), label = "country")
  col_name <- first_existing_col(df_raw, c("refinery_name", "refinery"), label = "refinery name")
  col_product <- first_existing_col(df_raw, c("product"), label = "product")
  col_grade <- first_existing_col(df_raw, c("grade"), label = "grade")
  col_volume <- first_existing_col(
    df_raw,
    c(
      "product_volume_b_d",
      "product_volume_bd",
      "product_volume_bbl_d",
      "product_volume_bbl_per_d",
      "volume_b_d",
      "product_volume"
    ),
    label = "product volume"
  )
  col_yield <- first_existing_col(
    df_raw,
    c("product_yield_vol_percent", "product_yield_vol", "yield_vol_percent", "yield_percent"),
    required = FALSE,
    label = "product yield"
  )
  
  out <- df_raw %>%
    transmute(
      region = if (!is.na(col_region)) as.character(.data[[col_region]]) else NA_character_,
      country_1 = as.character(.data[[col_country]]),
      rem_wm_refinery_name = as.character(.data[[col_name]]),
      product = as.character(.data[[col_product]]),
      grade = as.character(.data[[col_grade]]),
      rem_wm_hsfo_volume_bpd = safe_num(.data[[col_volume]]),
      product_yield_vol_perc = if (!is.na(col_yield)) safe_num(.data[[col_yield]]) else NA_real_
    )
  
  out
}

read_rem_hsfo_volumes <- function(files) {
  assert_files_exist(files)
  
  log_step("[3] Loading REM / WM regional refinery files")
  rem_all <- purrr::map_dfr(files, read_wm_sheet)
  
  rem_all %>%
    mutate(
      product = stringr::str_squish(product),
      grade = stringr::str_squish(grade),
      rem_wm_refinery_name = stringr::str_squish(rem_wm_refinery_name)
    ) %>%
    filter(
      stringr::str_to_lower(product) == "fuel oil",
      grade %in% c("High Sulphur", "Low Sulphur", "Very Low Sulphur"),
      !is.na(rem_wm_hsfo_volume_bpd),
      rem_wm_hsfo_volume_bpd > 0
    ) %>%
    mutate(name_norm = normalize_name(rem_wm_refinery_name))
}

read_rem_hsfo_prices <- function(file) {
  assert_file_exists(file)
  
  log_step("[4] Loading REM prices / export shares")
  df <- readr::read_csv(file, show_col_types = FALSE) %>%
    janitor::clean_names()
  
  required_base <- c("refinery", "product")
  if (!all(required_base %in% names(df))) {
    stop("REM prices CSV must contain at least: ", paste(required_base, collapse = ", "))
  }
  
  df <- df %>%
    filter(
      stringr::str_detect(product, regex("fuel oil", ignore_case = TRUE)),
      stringr::str_detect(product, regex("high sulphur|low sulphur|very low sulphur", ignore_case = TRUE)),
      product != "Feedstocks:High SulphurFuel Oil"
    ) %>%
    rename(
      rem_wm_refinery_name = refinery,
      product_orig_name = product
    ) %>%
    mutate(
      product = "Fuel Oil",
      grade = case_when(
        product_orig_name %in% c("Fuel Oil - High Sulphur", "High SulphurFuel Oil", "Fuel Oil-High Sulphur") ~ "High Sulphur",
        product_orig_name %in% c("Fuel Oil - Low Sulphur", "Low SulphurFuel Oil", "Fuel Oil-Low Sulphur") ~ "Low Sulphur",
        product_orig_name %in% c("Fuel Oil - Very Low Sulphur", "Very Low SulphurFuel Oil", "Fuel Oil-Very Low Sulphur") ~ "Very Low Sulphur",
        TRUE ~ NA_character_
      ),
      name_norm = normalize_name(rem_wm_refinery_name)
    )
  
  price_tonne_col <- pick_col_by_simplified_name(df, "^productprice(us|usd)?per(tonne|ton|t)$", label = "price USD/tonne")
  price_bbl_col <- pick_col_by_simplified_name(df, "^productprice(us|usd)?perbbl$", label = "price USD/bbl")
  local_share_col <- pick_col_by_simplified_name(df, "^localmarketshareper$", label = "local market share")
  export_share_col <- pick_col_by_simplified_name(df, "^exportmarketshareper$", label = "export market share")
  country_col <- first_existing_col(df, c("country"), label = "country")
  
  df %>%
    transmute(
      country_2 = as.character(.data[[country_col]]),
      rem_wm_refinery_name = as.character(rem_wm_refinery_name),
      product = as.character(product),
      grade = as.character(grade),
      product_orig_name = as.character(product_orig_name),
      product_price_usd_per_tonne = safe_num(.data[[price_tonne_col]]),
      product_price_usd_per_bbl = safe_num(.data[[price_bbl_col]]),
      local_market_share_per = safe_num(.data[[local_share_col]]),
      export_market_share_per = safe_num(.data[[export_share_col]]),
      name_norm = name_norm
    )
}

# ---------------------------------------------------------------------------
# 4) load inputs
# ---------------------------------------------------------------------------
prelim_hsfo <- read_prelim_hsfo(cfg$prelim_ci_xlsx)
prelim_loc <- read_prelim_locations(cfg$prelim_locations_xlsx, prelim_hsfo)
rem_hsfo_vol <- read_rem_hsfo_volumes(cfg$rem_files)
rem_hsfo_prices <- read_rem_hsfo_prices(cfg$rem_prices_csv)

log_step("[5] Merging REM volumes with prices / shares")
rem_hsfo <- rem_hsfo_vol %>%
  left_join(
    rem_hsfo_prices,
    by = c("name_norm", "product", "grade"),
    suffix = c("_vol", "_price")
  ) %>%
  mutate(
    rem_wm_refinery_name = dplyr::coalesce(rem_wm_refinery_name_vol, rem_wm_refinery_name_price),
    country_2 = dplyr::coalesce(country_2, country_1)
  ) %>%
  select(
    region,
    country_1,
    country_2,
    rem_wm_refinery_name,
    product,
    grade,
    product_orig_name,
    rem_wm_hsfo_volume_bpd,
    product_yield_vol_perc,
    product_price_usd_per_tonne,
    product_price_usd_per_bbl,
    local_market_share_per,
    export_market_share_per,
    name_norm
  )

# ---------------------------------------------------------------------------
# 5) coverage table and harmonised full join
# ---------------------------------------------------------------------------
log_step("[6] Reconciling PRELIM and REM refinery names")

prelim_name_table <- prelim_loc %>%
  transmute(
    prelim_refinery_name = as.character(prelim_refinery_name),
    name_norm = normalize_name(prelim_refinery_name)
  ) %>%
  filter(!is.na(name_norm), name_norm != "") %>%
  distinct()

rem_name_table <- rem_hsfo %>%
  transmute(
    rem_wm_refinery_name = as.character(rem_wm_refinery_name),
    name_norm = normalize_name(rem_wm_refinery_name)
  ) %>%
  filter(!is.na(name_norm), name_norm != "") %>%
  distinct()

coverage_status <- full_join(
  prelim_name_table %>% mutate(in_prelim = TRUE),
  rem_name_table %>% mutate(in_rem = TRUE),
  by = "name_norm"
) %>%
  mutate(
    display_name = dplyr::coalesce(prelim_refinery_name, rem_wm_refinery_name),
    status_wm_prelim = dplyr::case_when(
      !is.na(in_prelim) & !is.na(in_rem) ~ "both",
      is.na(in_prelim) & !is.na(in_rem) ~ "PRELIM_not_considering",
      !is.na(in_prelim) & is.na(in_rem) ~ "REM_not_considering",
      TRUE ~ "neither"
    )
  ) %>%
  select(name_norm, display_name, status_wm_prelim)

prelim_join <- prelim_loc %>%
  mutate(name_norm = normalize_name(prelim_refinery_name)) %>%
  select(
    refinery_id,
    prelim_refinery_name,
    country_prelim,
    lat_ref,
    lon_ref,
    throughput_volume_kbpd = any_of("throughput_volume_kbpd"),
    crude_intake_upstream_ci_kg_co2e_bbl = any_of("crude_intake_upstream_ci_kg_co2e_bbl"),
    crude_transport_ci_kg_co2e_bbl = any_of("crude_transport_ci_kg_co2e_bbl"),
    crude_intake_upstream_ci_g_co2e_mj = any_of("crude_intake_upstream_ci_g_co2e_mj"),
    crude_transport_ci_g_co2e_mj = any_of("crude_transport_ci_g_co2e_mj"),
    energy_density_lhv_mj_bbl = any_of("energy_density_lhv_mj_bbl"),
    hsfo_refining_ci_g_co2e_mj,
    hsfo_total_before_shipping,
    prelim_hsfo_volume_bpd,
    name_norm
  )

harmonised_all <- full_join(rem_hsfo, prelim_join, by = "name_norm") %>%
  left_join(coverage_status, by = "name_norm") %>%
  mutate(
    analysis_country = dplyr::coalesce(country_1, country_2, country_prelim),
    analysis_refinery_name = dplyr::coalesce(rem_wm_refinery_name, prelim_refinery_name, display_name),
    analysis_hsfo_volume_bpd = dplyr::coalesce(rem_wm_hsfo_volume_bpd, prelim_hsfo_volume_bpd)
  ) %>%
  relocate(
    analysis_refinery_name,
    analysis_country,
    status_wm_prelim,
    name_norm,
    .before = 1
  )

# ---------------------------------------------------------------------------
# 6) patch PRELIM_not_considering rows
# ---------------------------------------------------------------------------
log_step("[7] Patching PRELIM_not_considering rows")

harmonised_all <- harmonised_all %>%
  group_by(analysis_country) %>%
  mutate(
    prelim_refinery_name = if_else(
      status_wm_prelim == "PRELIM_not_considering" &
        (is.na(prelim_refinery_name) | prelim_refinery_name == ""),
      rem_wm_refinery_name,
      prelim_refinery_name
    ),
    prelim_hsfo_volume_bpd = if_else(
      status_wm_prelim == "PRELIM_not_considering" & is.na(prelim_hsfo_volume_bpd),
      rem_wm_hsfo_volume_bpd,
      prelim_hsfo_volume_bpd
    ),
    refinery_id = if_else(
      status_wm_prelim == "PRELIM_not_considering" &
        (is.na(refinery_id) | refinery_id == ""),
      paste0(
        "r_",
        dplyr::coalesce(
          suppressWarnings(countrycode(analysis_country, "country.name", "iso3c")),
          "UNK"
        ),
        "_",
        row_number()
      ),
      refinery_id
    )
  ) %>%
  ungroup() %>%
  mutate(
    analysis_hsfo_volume_bpd = dplyr::coalesce(rem_wm_hsfo_volume_bpd, prelim_hsfo_volume_bpd)
  )

# ---------------------------------------------------------------------------
# 7) country-level imputation
# ---------------------------------------------------------------------------
log_step("[8] Country-level imputations")

ci_cols <- intersect(
  c(
    "crude_intake_upstream_ci_kg_co2e_bbl",
    "crude_transport_ci_kg_co2e_bbl",
    "crude_intake_upstream_ci_g_co2e_mj",
    "crude_transport_ci_g_co2e_mj",
    "energy_density_lhv_mj_bbl",
    "hsfo_refining_ci_g_co2e_mj",
    "hsfo_total_before_shipping"
  ),
  names(harmonised_all)
)

price_cols <- intersect(
  c(
    "product_price_usd_per_tonne",
    "product_price_usd_per_bbl"
  ),
  names(harmonised_all)
)

share_cols <- intersect(
  c(
    "local_market_share_per",
    "export_market_share_per"
  ),
  names(harmonised_all)
)

country_ci <- harmonised_all %>%
  group_by(analysis_country) %>%
  summarise(
    across(
      all_of(ci_cols),
      ~ weighted_mean_na(.x, analysis_hsfo_volume_bpd),
      .names = "wavg_{.col}"
    ),
    .groups = "drop"
  )

country_price <- harmonised_all %>%
  group_by(analysis_country) %>%
  summarise(
    across(
      all_of(price_cols),
      mean_positive_na,
      .names = "avg_{.col}"
    ),
    .groups = "drop"
  )

country_share <- harmonised_all %>%
  group_by(analysis_country) %>%
  summarise(
    across(
      all_of(share_cols),
      mean_positive_na,
      .names = "avg_{.col}"
    ),
    .groups = "drop"
  )

harmonised_all <- harmonised_all %>%
  left_join(country_ci, by = "analysis_country") %>%
  left_join(country_price, by = "analysis_country") %>%
  left_join(country_share, by = "analysis_country")

if (length(ci_cols) > 0) {
  harmonised_all <- harmonised_all %>%
    mutate(
      across(
        all_of(ci_cols),
        ~ dplyr::coalesce(.x, .data[[paste0("wavg_", cur_column())]])
      )
    )
}

if (length(price_cols) > 0) {
  harmonised_all <- harmonised_all %>%
    mutate(
      across(
        all_of(price_cols),
        ~ dplyr::coalesce(.x, .data[[paste0("avg_", cur_column())]])
      )
    )
}

if (length(share_cols) > 0) {
  harmonised_all <- harmonised_all %>%
    mutate(
      across(
        all_of(share_cols),
        ~ dplyr::coalesce(.x, .data[[paste0("avg_", cur_column())]])
      )
    )
}

harmonised_all <- harmonised_all %>%
  select(
    -starts_with("wavg_"),
    -starts_with("avg_")
  )

# ---------------------------------------------------------------------------
# 8) region completion and explicit rules
# ---------------------------------------------------------------------------
log_step("[9] Filling missing regions")

harmonised_all <- harmonised_all %>%
  fill_region_from_country(region_col = "region", country_col = "country_1") %>%
  fill_region_from_country(region_col = "region", country_col = "country_2") %>%
  fill_region_from_country(region_col = "region", country_col = "analysis_country") %>%
  mutate(
    region = case_when(
      analysis_country == "Hungary" ~ "Europe",
      analysis_country == "Oman" ~ "Middle East",
      TRUE ~ region
    )
  )

# ---------------------------------------------------------------------------
# 9) final dataset variants
# ---------------------------------------------------------------------------
log_step("[10] Creating analysis-ready output")

harmonised_all <- harmonised_all %>%
  mutate(
    analysis_hsfo_volume_bpd = dplyr::coalesce(rem_wm_hsfo_volume_bpd, prelim_hsfo_volume_bpd)
  )

harmonised_analysis_ready <- harmonised_all %>%
  filter(
    !is.na(hsfo_refining_ci_g_co2e_mj),
    !is.na(analysis_hsfo_volume_bpd),
    analysis_hsfo_volume_bpd > 0
  )

# ---------------------------------------------------------------------------
# 10) diagnostics
# ---------------------------------------------------------------------------
log_step("[11] Building diagnostics")

status_counts <- harmonised_all %>%
  count(status_wm_prelim, name = "value") %>%
  mutate(metric = paste0("status_", status_wm_prelim)) %>%
  select(metric, value)

na_counts <- tibble(
  metric = c(
    "na_analysis_country",
    "na_region",
    "na_analysis_hsfo_volume_bpd",
    "na_hsfo_refining_ci_g_co2e_mj",
    "na_hsfo_total_before_shipping",
    "na_product_price_usd_per_tonne",
    "na_product_price_usd_per_bbl",
    "na_export_market_share_per",
    "na_lat_ref",
    "na_lon_ref"
  ),
  value = c(
    sum(is.na(harmonised_all$analysis_country)),
    sum(is.na(harmonised_all$region)),
    sum(is.na(harmonised_all$analysis_hsfo_volume_bpd)),
    sum(is.na(harmonised_all$hsfo_refining_ci_g_co2e_mj)),
    sum(is.na(harmonised_all$hsfo_total_before_shipping)),
    sum(is.na(harmonised_all$product_price_usd_per_tonne)),
    sum(is.na(harmonised_all$product_price_usd_per_bbl)),
    sum(is.na(harmonised_all$export_market_share_per)),
    sum(is.na(harmonised_all$lat_ref)),
    sum(is.na(harmonised_all$lon_ref))
  )
)

summary_metrics <- tibble(
  metric = c(
    "rows_prelim_hsfo",
    "rows_prelim_locations_joined",
    "rows_rem_hsfo_volumes",
    "rows_rem_hsfo_prices",
    "rows_harmonised_all",
    "rows_harmonised_analysis_ready",
    "unique_refineries_harmonised_all",
    "unique_refineries_analysis_ready",
    "total_rem_hsfo_volume_bpd",
    "total_prelim_hsfo_volume_bpd",
    "total_analysis_hsfo_volume_bpd"
  ),
  value = c(
    nrow(prelim_hsfo),
    nrow(prelim_loc),
    nrow(rem_hsfo_vol),
    nrow(rem_hsfo_prices),
    nrow(harmonised_all),
    nrow(harmonised_analysis_ready),
    dplyr::n_distinct(harmonised_all$analysis_refinery_name),
    dplyr::n_distinct(harmonised_analysis_ready$analysis_refinery_name),
    sum(harmonised_all$rem_wm_hsfo_volume_bpd, na.rm = TRUE),
    sum(harmonised_all$prelim_hsfo_volume_bpd, na.rm = TRUE),
    sum(harmonised_all$analysis_hsfo_volume_bpd, na.rm = TRUE)
  )
)

diagnostics <- bind_rows(summary_metrics, status_counts, na_counts)

# ---------------------------------------------------------------------------
# 11) write outputs
# ---------------------------------------------------------------------------
log_step("[12] Writing outputs")

ts_tag <- format(Sys.time(), "%Y%m%d_%H%M")

readr::write_csv(
  prelim_hsfo,
  file.path(cfg$out_dir, paste0("01_prelim_hsfo_asset_level_", ts_tag, ".csv"))
)

readr::write_csv(
  prelim_loc,
  file.path(cfg$out_dir, paste0("02_prelim_hsfo_with_locations_", ts_tag, ".csv"))
)

readr::write_csv(
  rem_hsfo_vol,
  file.path(cfg$out_dir, paste0("03_rem_hsfo_volumes_", ts_tag, ".csv"))
)

readr::write_csv(
  rem_hsfo_prices,
  file.path(cfg$out_dir, paste0("04_rem_hsfo_prices_shares_", ts_tag, ".csv"))
)

readr::write_csv(
  harmonised_all,
  file.path(cfg$out_dir, paste0("05_hsfo_harmonised_all_rows_", ts_tag, ".csv"))
)

readr::write_csv(
  harmonised_analysis_ready,
  file.path(cfg$out_dir, paste0("06_hsfo_harmonised_analysis_ready_", ts_tag, ".csv"))
)

readr::write_csv(
  diagnostics,
  file.path(cfg$out_dir, paste0("07_hsfo_harmonisation_diagnostics_", ts_tag, ".csv"))
)

writeLines(
  capture.output(sessionInfo()),
  file.path(cfg$out_dir, paste0("08_session_info_", ts_tag, ".txt"))
)

log_step("Done")
message("Outputs written to: ", normalizePath(cfg$out_dir))





###############################################################################
# LPG well-to-refinery harmonisation workflow
#
# Purpose:
# Extract LPG asset-level carbon-intensity outputs from a PRELIM workbook,
# optionally attach refinery coordinates, and write analysis-ready tables.
#
# Scope:
# - no plots
# - reproducible and repository-ready
# - limited to PRELIM LPG outputs and refinery locations
#
# Main outputs:
# 01_lpg_well_to_refinery_ci_asset_level_<timestamp>.csv
# 02_lpg_well_to_refinery_ci_with_locations_<timestamp>.csv
# 03_lpg_diagnostics_<timestamp>.csv
# 04_session_info_<timestamp>.txt
#
# Notes:
# - this script does not run PRELIM
# - this script only post-processes PRELIM outputs
###############################################################################

rm(list = ls())

# ---------------------------------------------------------------------------
# 0) required packages
# ---------------------------------------------------------------------------
required_pkgs <- c(
  "dplyr",
  "readr",
  "readxl",
  "janitor",
  "tibble"
)

missing_pkgs <- setdiff(required_pkgs, rownames(installed.packages()))
if (length(missing_pkgs) > 0) {
  stop(
    "Install missing packages before running this script: ",
    paste(missing_pkgs, collapse = ", ")
  )
}

invisible(lapply(required_pkgs, library, character.only = TRUE))

# ---------------------------------------------------------------------------
# 1) user configuration
# ---------------------------------------------------------------------------
cfg <- list(
  prelim_ci_xlsx = "PATH/Marine fuel well to refinery exit CI v5 (PRELIMv20) LPG lower v2 2023.xlsx",
  refinery_locations_xlsx = "PATH/refinery locations.xlsx",
  out_dir = "PATH/04 LPG WTRef"
)

dir.create(cfg$out_dir, showWarnings = FALSE, recursive = TRUE)

# ---------------------------------------------------------------------------
# 2) helper functions
# ---------------------------------------------------------------------------
log_step <- function(x) message("\n", x)

assert_file_exists <- function(path) {
  if (!file.exists(path)) stop("File not found: ", path)
}

safe_num <- function(x) suppressWarnings(as.numeric(x))

first_existing_col <- function(df, candidates, required = TRUE, label = "column") {
  hit <- intersect(candidates, names(df))
  if (length(hit) == 0) {
    if (required) {
      stop(
        "No ", label, " found. Candidates tested: ",
        paste(candidates, collapse = ", "),
        "\nAvailable columns: ",
        paste(names(df), collapse = ", ")
      )
    } else {
      return(NA_character_)
    }
  }
  hit[1]
}

# ---------------------------------------------------------------------------
# 3) read and clean PRELIM data
# ---------------------------------------------------------------------------
read_prelim_lpg <- function(file) {
  assert_file_exists(file)
  
  log_step("[1] Loading PRELIM workbook")
  df <- readxl::read_excel(file, sheet = 1) %>%
    janitor::clean_names() %>%
    as.data.frame()
  
  # Standardise key names where needed
  if ("x1" %in% names(df)) names(df)[names(df) == "x1"] <- "refinery_id"
  if ("mj_bbl" %in% names(df)) names(df)[names(df) == "mj_bbl"] <- "energy_density_lhv_mj_bbl"
  
  # Remove known unused columns when present
  drop_cols <- c(
    "x8",
    "x11",
    "x18",
    "hsfo_domestic_use_bbl_d",
    "hsfo_import_bbl_d",
    "fuel_trade_shipping_ci_g_co2e_mj",
    "combustion_ci_g_co2e_mj"
  )
  
  df <- df %>%
    select(-any_of(drop_cols))
  
  # Add a refinery name field if available
  refinery_name_col <- first_existing_col(
    df,
    candidates = c("prelim_refinery_name", "refinery_name", "refinery", "name"),
    required = FALSE,
    label = "PRELIM refinery name"
  )
  
  if (is.na(refinery_name_col)) {
    df$prelim_refinery_name <- NA_character_
  } else {
    df$prelim_refinery_name <- as.character(df[[refinery_name_col]])
  }
  
  # Validate LPG columns
  needed <- c(
    "refinery_id",
    "lpg_volume_bbl_d",
    "lpg_refining_ci_g_co2e_mj",
    "lpg_total_before_shipping"
  )
  
  missing_needed <- setdiff(needed, names(df))
  if (length(missing_needed) > 0) {
    stop(
      "PRELIM workbook is missing required LPG columns: ",
      paste(missing_needed, collapse = ", ")
    )
  }
  
  # Keep LPG-relevant content only
  lpg <- df %>%
    select(-any_of(c(
      "hsfo_volume_bbl_d",
      "hsfo_refining_ci_g_co2e_mj",
      "hsfo_total_before_shipping"
    ))) %>%
    mutate(
      lpg_volume_bbl_d = safe_num(lpg_volume_bbl_d),
      lpg_refining_ci_g_co2e_mj = safe_num(lpg_refining_ci_g_co2e_mj),
      lpg_total_before_shipping = safe_num(lpg_total_before_shipping)
    ) %>%
    filter(
      !is.na(lpg_volume_bbl_d),
      !is.na(lpg_refining_ci_g_co2e_mj),
      !is.na(lpg_total_before_shipping)
    )
  
  lpg
}

# ---------------------------------------------------------------------------
# 4) attach refinery locations
# ---------------------------------------------------------------------------
attach_locations <- function(lpg_df, locations_file) {
  assert_file_exists(locations_file)
  
  log_step("[2] Loading refinery locations")
  loc <- readxl::read_excel(locations_file, sheet = 1) %>%
    janitor::clean_names() %>%
    as.data.frame()
  
  id_col <- first_existing_col(loc, c("refinery_id"), label = "refinery_id")
  lat_col <- first_existing_col(loc, c("latitude", "lat"), label = "latitude")
  lon_col <- first_existing_col(loc, c("longitude", "lon", "lng"), label = "longitude")
  
  loc <- loc %>%
    rename(
      refinery_id = !!id_col,
      lat_ref = !!lat_col,
      lon_ref = !!lon_col
    ) %>%
    select(refinery_id, lat_ref, lon_ref)
  
  lpg_df %>%
    left_join(loc, by = "refinery_id")
}

# ---------------------------------------------------------------------------
# 5) run workflow
# ---------------------------------------------------------------------------
lpg_asset_level <- read_prelim_lpg(cfg$prelim_ci_xlsx)

log_step("[3] Attaching coordinates to LPG dataset")
lpg_with_locations <- attach_locations(lpg_asset_level, cfg$refinery_locations_xlsx)

# ---------------------------------------------------------------------------
# 6) diagnostics
# ---------------------------------------------------------------------------
log_step("[4] Building diagnostics")

diagnostics <- tibble::tibble(
  metric = c(
    "rows_lpg_asset_level",
    "rows_lpg_with_locations",
    "unique_refinery_id_asset_level",
    "unique_refinery_name_asset_level",
    "sum_lpg_volume_bbl_d_asset_level",
    "rows_missing_lat",
    "rows_missing_lon",
    "rows_missing_any_coordinates"
  ),
  value = c(
    nrow(lpg_asset_level),
    nrow(lpg_with_locations),
    dplyr::n_distinct(lpg_asset_level$refinery_id),
    if ("prelim_refinery_name" %in% names(lpg_asset_level)) {
      dplyr::n_distinct(lpg_asset_level$prelim_refinery_name, na.rm = TRUE)
    } else {
      NA_integer_
    },
    sum(lpg_asset_level$lpg_volume_bbl_d, na.rm = TRUE),
    sum(is.na(lpg_with_locations$lat_ref)),
    sum(is.na(lpg_with_locations$lon_ref)),
    sum(is.na(lpg_with_locations$lat_ref) | is.na(lpg_with_locations$lon_ref))
  )
)

# ---------------------------------------------------------------------------
# 7) write outputs
# ---------------------------------------------------------------------------
log_step("[5] Writing outputs")

ts_tag <- format(Sys.time(), "%Y%m%d_%H%M")

readr::write_csv(
  lpg_asset_level,
  file.path(cfg$out_dir, paste0("01_lpg_well_to_refinery_ci_asset_level_", ts_tag, ".csv"))
)

readr::write_csv(
  lpg_with_locations,
  file.path(cfg$out_dir, paste0("02_lpg_well_to_refinery_ci_with_locations_", ts_tag, ".csv"))
)

readr::write_csv(
  diagnostics,
  file.path(cfg$out_dir, paste0("03_lpg_diagnostics_", ts_tag, ".csv"))
)

writeLines(
  capture.output(sessionInfo()),
  file.path(cfg$out_dir, paste0("04_session_info_", ts_tag, ".txt"))
)

log_step("Done")
message("Outputs written to: ", normalizePath(cfg$out_dir))



###############################################################################
# LNG value-chain harmonisation and CI calculation workflow
#
# Purpose:
# Build analysis-ready LNG terminal, regas terminal, and route-level datasets
# by combining:
#   1) Archie LNG spatial layers
#   2) updated LNG liquefaction results
#   3) route-level joins from liquefaction terminals to regas terminals
#
# Outputs:
# 01_lng_terminals_updated_<timestamp>.csv
# 02_regas_terminals_clean_<timestamp>.csv
# 03_lng_sea_routes_clean_<timestamp>.csv
# 04_lng_routes_well_to_regas_<timestamp>.csv
# 05_unmatched_lng_origins_<timestamp>.csv
# 06_unmatched_regas_destinations_<timestamp>.csv
# 07_lng_harmonisation_diagnostics_<timestamp>.csv
# 08_session_info_<timestamp>.txt
#
# Source attribution:
# Base LNG dataset derived from the Research Square preprint:
# "Global carbon intensity of LNG value chain amid intensifying energy security
# and climate trade-offs", DOI: 10.21203/rs.3.rs-7255872/v1
# License: CC BY 4.0
#
# Important assumption:
# The column specified in cfg$updated_liquefaction_ci_col is treated as the
# updated liquefaction-stage CI in gCO2e/MJ. If that column instead represents
# a broader boundary, adjust the aggregation logic below before use.
###############################################################################

rm(list = ls())

# ---------------------------------------------------------------------------
# 0) required packages
# ---------------------------------------------------------------------------
required_pkgs <- c(
  "dplyr",
  "sf",
  "readr",
  "stringr",
  "tibble",
  "countrycode"
)

missing_pkgs <- setdiff(required_pkgs, rownames(installed.packages()))
if (length(missing_pkgs) > 0) {
  stop(
    "Install missing packages before running this script: ",
    paste(missing_pkgs, collapse = ", ")
  )
}

invisible(lapply(required_pkgs, library, character.only = TRUE))

# ---------------------------------------------------------------------------
# 1) user configuration
# ---------------------------------------------------------------------------
cfg <- list(
  archie_root = "PATH/Archie v15",
  updated_lng_csv = "PATH/LNG_CI_calculations_MERGED_UPDATED.csv",
  out_dir = "PATH/07 repository_ready_outputs",
  
  lng_terminals_file = "lngs.geojson",
  sea_routes_file = "gasSeaRoutes-with-meta.geojson",
  regas_terminals_file = "regas-combined.geojson",
  
  # Column in the updated CSV interpreted as updated liquefaction CI
  updated_liquefaction_ci_col = "Final_LNG_CI_gCO2e_per_MJ"
)

dir.create(cfg$out_dir, showWarnings = FALSE, recursive = TRUE)

# ---------------------------------------------------------------------------
# 2) helper functions
# ---------------------------------------------------------------------------
log_step <- function(x) message("\n", x)

assert_file_exists <- function(path) {
  if (!file.exists(path)) stop("File not found: ", path)
}

safe_num <- function(x) suppressWarnings(as.numeric(x))

normalize_key <- function(x) {
  x %>%
    as.character() %>%
    tolower() %>%
    stringr::str_replace_all("[^a-z0-9]+", " ") %>%
    stringr::str_squish()
}

first_existing_col <- function(df, candidates, required = TRUE, label = "column") {
  hit <- intersect(candidates, names(df))
  if (length(hit) == 0) {
    if (required) {
      stop(
        "No ", label, " found. Candidates tested: ",
        paste(candidates, collapse = ", "),
        "\nAvailable columns: ",
        paste(names(df), collapse = ", ")
      )
    } else {
      return(NA_character_)
    }
  }
  hit[1]
}

read_sf_layer <- function(path) {
  assert_file_exists(path)
  x <- sf::st_read(path, quiet = TRUE)
  if (is.na(sf::st_crs(x))) sf::st_crs(x) <- 4326
  x
}

# Deduplicate by key while preserving one record per key
# This avoids row multiplication during joins.
dedupe_by_key <- function(df, key_col) {
  key_sym <- rlang::sym(key_col)
  df %>%
    group_by(!!key_sym) %>%
    slice(1) %>%
    ungroup()
}

# ---------------------------------------------------------------------------
# 3) file paths
# ---------------------------------------------------------------------------
lng_terminals_path <- file.path(cfg$archie_root, cfg$lng_terminals_file)
sea_routes_path <- file.path(cfg$archie_root, cfg$sea_routes_file)
regas_terminals_path <- file.path(cfg$archie_root, cfg$regas_terminals_file)

assert_file_exists(lng_terminals_path)
assert_file_exists(sea_routes_path)
assert_file_exists(regas_terminals_path)
assert_file_exists(cfg$updated_lng_csv)

# ---------------------------------------------------------------------------
# 4) name harmonisation map: Archie LNG names
# ---------------------------------------------------------------------------
name_map <- c(
  "Bethioua"               = "Arzew-Bethioua",
  "Skikda"                 = "Skikda",
  "APLNG"                  = "APLNG",
  "Darwin"                 = "Darwin",
  "GLNG"                   = "Gladstone",
  "Gorgon"                 = "Gorgon",
  "Ichthys"                = "Ichthys",
  "NWS"                    = "North West Shelf",
  "Pluto"                  = "Pluto",
  "Prelude"                = "Prelude FLNG",
  "QCLNG"                  = "QCLNG",
  "Wheatstone"             = "Wheatstone",
  "Bontang"                = "Bontang",
  "DSLNG"                  = "DSLNG",
  "Tangguh"                = "Tangguh",
  "Bintulu"                = "Bintulu",
  "PFLNG 1 Sabah"          = "PFLNG 1 Sabah",
  "PFLNG 2"                = "PFLNG 2",
  "Bonny LNG"              = "Bonny",
  "Qalhat"                 = "Qalhat",
  "PNG LNG"                = "PNG LNG",
  "Ras Laffan"             = "Ras Laffan",
  "Sakhalin 2"             = "Sakhalin II",
  "Yamal"                  = "Yamal",
  "Portovaya"              = "Portovaya",
  "Vysotsk"                = "Vysotsk",
  "Calcasieu Pass"         = "Calcasieu Pass",
  "Cameron (Liqu.)"        = "Cameron",
  "Corpus Christi"         = "Corpus Christi",
  "Cove Point"             = "Cove Point",
  "Elba Island Liq."       = "Elba Island",
  "Freeport"               = "Freeport Texas",
  "Sabine Pass"            = "Sabine Pass",
  "Das Island"             = "Das Island",
  "Atlantic LNG"           = "Atlantic LNG",
  "Idku"                   = "Idku",
  "Peru LNG"               = "Peru LNG",
  "Damietta"               = "Damietta",
  "Cameroon FLNG Terminal" = "Cameroon FLNG Terminal",
  "Bioko"                  = "Bioko",
  "Soyo"                   = "Soyo",
  "Snohvit"                = "Snohvit",
  "Lumut I"                = "Lumut I"
)

# ---------------------------------------------------------------------------
# 5) updated LNG liquefaction results
# ---------------------------------------------------------------------------
log_step("[1] Loading updated LNG liquefaction results")

lng_updated_raw <- readr::read_csv(cfg$updated_lng_csv, show_col_types = FALSE)

updated_name_col <- first_existing_col(
  lng_updated_raw,
  candidates = c("lng_archie", "lng_name", "LNG.Name", "LNG_Name"),
  label = "updated LNG terminal name"
)

updated_ci_col <- first_existing_col(
  lng_updated_raw,
  candidates = c(
    cfg$updated_liquefaction_ci_col,
    "final_lng_ci_gco2e_per_mj",
    "Final_LNG_CI_gCO2e_per_MJ"
  ),
  label = "updated liquefaction CI column"
)

lng_updated <- lng_updated_raw %>%
  transmute(
    lng_archie = stringr::str_squish(as.character(.data[[updated_name_col]])),
    liquefaction_gco2e_mj_updated = safe_num(.data[[updated_ci_col]])
  ) %>%
  filter(!is.na(lng_archie), lng_archie != "") %>%
  distinct(lng_archie, .keep_all = TRUE)

# ---------------------------------------------------------------------------
# 6) LNG terminals
# ---------------------------------------------------------------------------
log_step("[2] Loading LNG terminal layer")

lngs_raw <- read_sf_layer(lng_terminals_path)

lng_id_col <- first_existing_col(lngs_raw, c("LNG.ID"), label = "LNG terminal ID")
lng_name_col <- first_existing_col(lngs_raw, c("LNG.Name"), label = "LNG terminal name")
lng_country_col <- first_existing_col(lngs_raw, c("country", "Country"), required = FALSE, label = "LNG country")
lng_iso_col <- first_existing_col(lngs_raw, c("Country.ISO.code"), required = FALSE, label = "LNG ISO code")
lng_export_col <- first_existing_col(lngs_raw, c("Exporting.LNG..tonne.year"), required = FALSE, label = "export tonnes/year")

lngs_base <- lngs_raw %>%
  sf::st_drop_geometry() %>%
  transmute(
    lng_id = safe_num(.data[[lng_id_col]]),
    lng_name = stringr::str_squish(as.character(.data[[lng_name_col]])),
    country = if (!is.na(lng_country_col)) as.character(.data[[lng_country_col]]) else NA_character_,
    iso3 = if (!is.na(lng_iso_col)) as.character(.data[[lng_iso_col]]) else NA_character_,
    export_tpy = if (!is.na(lng_export_col)) safe_num(.data[[lng_export_col]]) else NA_real_,
    
    surface_processing_gco2e_mj    = safe_num(.data[["Surface.Processing..gCO2eq..MJ"]]),
    flaring_gco2e_mj               = safe_num(.data[["Flaring..gCO2eq..MJ"]]),
    co2_venting_agr_gco2e_mj       = safe_num(.data[["CO2.Venting...AGR..gCO2eq..MJ"]]),
    production_extraction_gco2e_mj = safe_num(.data[["Production.and.Extraction..gCO2eq..MJ"]]),
    methane_leak_super_gco2e_mj    = safe_num(.data[["Methane.Leak...Super.Emitter...Ave..gCO2eq..MJ"]]),
    methane_leak_other_gco2e_mj    = safe_num(.data[["Methane.Leak...Other...Ave..gCO2eq..MJ"]]),
    exploration_drilling_gco2e_mj  = safe_num(.data[["Exploration.and.Drilling..gCO2eq..MJ"]]),
    offsite_gco2e_mj               = safe_num(.data[["Offsite..gCO2eq..MJ"]]),
    crude_oil_transport_gco2e_mj   = safe_num(.data[["Crude.Oil.Transportation..gCO2eq..MJ"]]),
    f2pp_transmission_gco2e_mj     = safe_num(.data[["F2PP.Transmission..gCO2eq..MJ"]]),
    pp2l_transmission_gco2e_mj     = safe_num(.data[["PP2L.Transmission..gCO2eq..MJ"]]),
    liquefaction_gco2e_mj_base     = safe_num(.data[["Liquefaction..gCO2eq..MJ"]])
  ) %>%
  mutate(
    lng_archie = dplyr::recode(lng_name, !!!name_map, .default = lng_name),
    iso3 = dplyr::coalesce(
      iso3,
      suppressWarnings(countrycode::countrycode(country, "country.name", "iso3c"))
    )
  )

lng_terminals <- lngs_base %>%
  left_join(lng_updated, by = "lng_archie") %>%
  mutate(
    liquefaction_gco2e_mj = dplyr::coalesce(liquefaction_gco2e_mj_updated, liquefaction_gco2e_mj_base),
    liquefaction_ci_source = dplyr::case_when(
      !is.na(liquefaction_gco2e_mj_updated) ~ "updated_csv",
      !is.na(liquefaction_gco2e_mj_base) ~ "base_geojson",
      TRUE ~ NA_character_
    ),
    origin_key = normalize_key(lng_name)
  )

# ---------------------------------------------------------------------------
# 7) regas terminals
# ---------------------------------------------------------------------------
log_step("[3] Loading regas terminal layer")

regas_raw <- read_sf_layer(regas_terminals_path)

regas_id_col <- first_existing_col(regas_raw, c("Regas.ID"), label = "regas ID")
regas_name_col <- first_existing_col(regas_raw, c("Regas.Name"), label = "regas name")
regas_country_col <- first_existing_col(regas_raw, c("Regas.Country"), required = FALSE, label = "regas country")
regas_iso_col <- first_existing_col(regas_raw, c("Regas.Country.ISO"), required = FALSE, label = "regas ISO")
regas_import_col <- first_existing_col(regas_raw, c("Importing.LNG..tonne.year"), required = FALSE, label = "import tonnes/year")
ship_ci_col <- first_existing_col(regas_raw, c("Shipping..gCO2eq..MJ"), required = FALSE, label = "shipping CI")
regas_ci_col <- first_existing_col(regas_raw, c("Regasification..gCO2eq..MJ"), required = FALSE, label = "regasification CI")

regas_terminals <- regas_raw %>%
  sf::st_drop_geometry() %>%
  transmute(
    regas_id = safe_num(.data[[regas_id_col]]),
    regas_name = stringr::str_squish(as.character(.data[[regas_name_col]])),
    regas_country = if (!is.na(regas_country_col)) as.character(.data[[regas_country_col]]) else NA_character_,
    regas_country_iso = if (!is.na(regas_iso_col)) as.character(.data[[regas_iso_col]]) else NA_character_,
    import_tpy = if (!is.na(regas_import_col)) safe_num(.data[[regas_import_col]]) else NA_real_,
    ship_ci_gco2e_mj = if (!is.na(ship_ci_col)) safe_num(.data[[ship_ci_col]]) else NA_real_,
    regas_ci_gco2e_mj = if (!is.na(regas_ci_col)) safe_num(.data[[regas_ci_col]]) else NA_real_
  ) %>%
  mutate(
    destination_key = normalize_key(regas_name)
  )

# ---------------------------------------------------------------------------
# 8) sea routes
# ---------------------------------------------------------------------------
log_step("[4] Loading LNG sea routes")

routes_raw <- read_sf_layer(sea_routes_path)

routes <- routes_raw %>%
  sf::st_drop_geometry() %>%
  transmute(
    origin = as.character(origin),
    destination = as.character(destination),
    origin_iso_code = dplyr::if_else("origin_iso_code" %in% names(routes_raw), as.character(origin_iso_code), NA_character_),
    destination_iso_code = dplyr::if_else("destination_iso_code" %in% names(routes_raw), as.character(destination_iso_code), NA_character_),
    archie_origin_terminal_id = dplyr::if_else("archie_origin_terminal_id" %in% names(routes_raw), as.character(archie_origin_terminal_id), NA_character_),
    archie_destination_terminal_id = dplyr::if_else("archie_destination_terminal_id" %in% names(routes_raw), as.character(archie_destination_terminal_id), NA_character_),
    route_name = dplyr::if_else("name" %in% names(routes_raw), as.character(name), NA_character_),
    route_desc = dplyr::if_else("Description" %in% names(routes_raw), as.character(Description), NA_character_),
    route_geojson_id = dplyr::if_else("route" %in% names(routes_raw), as.character(route), NA_character_)
  ) %>%
  mutate(
    origin = stringr::str_squish(origin),
    destination = stringr::str_squish(destination),
    origin_key = normalize_key(origin),
    destination_key = normalize_key(destination)
  )

# ---------------------------------------------------------------------------
# 9) join routes to LNG terminals and regas terminals
# ---------------------------------------------------------------------------
log_step("[5] Joining routes to LNG and regas terminals")

lng_terminals_unique <- lng_terminals %>%
  dedupe_by_key("origin_key")

regas_terminals_unique <- regas_terminals %>%
  dedupe_by_key("destination_key")

routes_lng <- routes %>%
  left_join(
    lng_terminals_unique %>%
      select(
        -origin_key,
        lng_name
      ),
    by = c("origin_key")
  ) %>%
  mutate(
    lng_match_missing = is.na(lng_id)
  )

routes_lng_regas <- routes_lng %>%
  left_join(
    regas_terminals_unique %>%
      select(
        -destination_key,
        regas_name
      ),
    by = c("destination_key")
  ) %>%
  mutate(
    regas_match_missing = is.na(regas_id)
  )

# ---------------------------------------------------------------------------
# 10) route-level CI calculations
# ---------------------------------------------------------------------------
log_step("[6] Calculating route-level CI metrics")

component_cols_wtl <- c(
  "surface_processing_gco2e_mj",
  "flaring_gco2e_mj",
  "co2_venting_agr_gco2e_mj",
  "production_extraction_gco2e_mj",
  "methane_leak_super_gco2e_mj",
  "methane_leak_other_gco2e_mj",
  "exploration_drilling_gco2e_mj",
  "offsite_gco2e_mj",
  "f2pp_transmission_gco2e_mj",
  "pp2l_transmission_gco2e_mj",
  "liquefaction_gco2e_mj"
)

well_to_regas <- routes_lng_regas %>%
  mutate(
    across(all_of(component_cols_wtl), safe_num),
    ship_ci_gco2e_mj = safe_num(ship_ci_gco2e_mj),
    regas_ci_gco2e_mj = safe_num(regas_ci_gco2e_mj)
  ) %>%
  mutate(
    wtl_ci_gco2e_mj = rowSums(across(all_of(component_cols_wtl)), na.rm = TRUE),
    wts_ci_gco2e_mj = dplyr::if_else(
      is.na(ship_ci_gco2e_mj),
      NA_real_,
      wtl_ci_gco2e_mj + ship_ci_gco2e_mj
    ),
    wtr_ci_gco2e_mj = dplyr::if_else(
      is.na(ship_ci_gco2e_mj) | is.na(regas_ci_gco2e_mj),
      NA_real_,
      wtl_ci_gco2e_mj + ship_ci_gco2e_mj + regas_ci_gco2e_mj
    )
  )

# ---------------------------------------------------------------------------
# 11) unmatched diagnostics
# ---------------------------------------------------------------------------
log_step("[7] Building match diagnostics")

unmatched_lng_origins <- well_to_regas %>%
  filter(lng_match_missing) %>%
  distinct(origin) %>%
  arrange(origin)

unmatched_regas_destinations <- well_to_regas %>%
  filter(regas_match_missing) %>%
  distinct(destination) %>%
  arrange(destination)

lng_key_duplicates <- lng_terminals %>%
  count(origin_key, name = "n") %>%
  filter(n > 1)

regas_key_duplicates <- regas_terminals %>%
  count(destination_key, name = "n") %>%
  filter(n > 1)

diagnostics <- tibble::tibble(
  metric = c(
    "rows_lng_terminals",
    "rows_regas_terminals",
    "rows_sea_routes",
    "rows_final_routes",
    "distinct_lng_names",
    "distinct_regas_names",
    "distinct_route_origins",
    "distinct_route_destinations",
    "unmatched_lng_origins",
    "unmatched_regas_destinations",
    "duplicate_lng_keys",
    "duplicate_regas_keys",
    "routes_missing_lng_match",
    "routes_missing_regas_match",
    "sum_export_tpy_lng",
    "sum_import_tpy_regas",
    "wtl_ci_min",
    "wtl_ci_mean",
    "wtl_ci_max",
    "wts_ci_min",
    "wts_ci_mean",
    "wts_ci_max",
    "wtr_ci_min",
    "wtr_ci_mean",
    "wtr_ci_max"
  ),
  value = c(
    nrow(lng_terminals),
    nrow(regas_terminals),
    nrow(routes),
    nrow(well_to_regas),
    dplyr::n_distinct(lng_terminals$lng_name),
    dplyr::n_distinct(regas_terminals$regas_name),
    dplyr::n_distinct(routes$origin),
    dplyr::n_distinct(routes$destination),
    nrow(unmatched_lng_origins),
    nrow(unmatched_regas_destinations),
    nrow(lng_key_duplicates),
    nrow(regas_key_duplicates),
    sum(well_to_regas$lng_match_missing, na.rm = TRUE),
    sum(well_to_regas$regas_match_missing, na.rm = TRUE),
    sum(lng_terminals$export_tpy, na.rm = TRUE),
    sum(regas_terminals$import_tpy, na.rm = TRUE),
    min(well_to_regas$wtl_ci_gco2e_mj, na.rm = TRUE),
    mean(well_to_regas$wtl_ci_gco2e_mj, na.rm = TRUE),
    max(well_to_regas$wtl_ci_gco2e_mj, na.rm = TRUE),
    min(well_to_regas$wts_ci_gco2e_mj, na.rm = TRUE),
    mean(well_to_regas$wts_ci_gco2e_mj, na.rm = TRUE),
    max(well_to_regas$wts_ci_gco2e_mj, na.rm = TRUE),
    min(well_to_regas$wtr_ci_gco2e_mj, na.rm = TRUE),
    mean(well_to_regas$wtr_ci_gco2e_mj, na.rm = TRUE),
    max(well_to_regas$wtr_ci_gco2e_mj, na.rm = TRUE)
  )
)

# ---------------------------------------------------------------------------
# 12) write outputs
# ---------------------------------------------------------------------------
log_step("[8] Writing outputs")

ts_tag <- format(Sys.time(), "%Y%m%d_%H%M")

readr::write_csv(
  lng_terminals,
  file.path(cfg$out_dir, paste0("01_lng_terminals_updated_", ts_tag, ".csv"))
)

readr::write_csv(
  regas_terminals,
  file.path(cfg$out_dir, paste0("02_regas_terminals_clean_", ts_tag, ".csv"))
)

readr::write_csv(
  routes,
  file.path(cfg$out_dir, paste0("03_lng_sea_routes_clean_", ts_tag, ".csv"))
)

readr::write_csv(
  well_to_regas,
  file.path(cfg$out_dir, paste0("04_lng_routes_well_to_regas_", ts_tag, ".csv"))
)

readr::write_csv(
  unmatched_lng_origins,
  file.path(cfg$out_dir, paste0("05_unmatched_lng_origins_", ts_tag, ".csv"))
)

readr::write_csv(
  unmatched_regas_destinations,
  file.path(cfg$out_dir, paste0("06_unmatched_regas_destinations_", ts_tag, ".csv"))
)

readr::write_csv(
  diagnostics,
  file.path(cfg$out_dir, paste0("07_lng_harmonisation_diagnostics_", ts_tag, ".csv"))
)

writeLines(
  capture.output(sessionInfo()),
  file.path(cfg$out_dir, paste0("08_session_info_", ts_tag, ".txt"))
)

log_step("Done")
message("Outputs written to: ", normalizePath(cfg$out_dir))
