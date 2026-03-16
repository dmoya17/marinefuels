###############################################################################
# Shared helpers for marine fuels figures
###############################################################################

suppressPackageStartupMessages({
  library(dplyr)
  library(readr)
  library(readxl)
  library(stringr)
  library(tidyr)
  library(ggplot2)
  library(ggsci)
  library(scales)
  library(grid)
  library(gridExtra)
  library(openxlsx)
  library(countrycode)
  library(sf)
  library(patchwork)
  library(tibble)
})

if (!exists(".pt")) .pt <- 72.27 / 25.4

BASE <- 13

pt_size <- function(x) x / .pt

nz_num <- function(x, repl = 0) {
  x <- suppressWarnings(as.numeric(x))
  ifelse(is.na(x), repl, x)
}

safe_div <- function(num, den) {
  ifelse(is.na(den) | den == 0, NA_real_, num / den)
}

wmean_na <- function(x, w) {
  x <- suppressWarnings(as.numeric(x))
  w <- suppressWarnings(as.numeric(w))
  ok <- !is.na(x) & !is.na(w) & is.finite(x) & is.finite(w) & w > 0
  if (!any(ok)) return(NA_real_)
  sum(x[ok] * w[ok]) / sum(w[ok])
}

safe_sheet_names <- function(x) {
  x <- gsub("[:?*/\\[\\]]", "_", x)
  x <- ifelse(nchar(x) > 31, substr(x, 1, 31), x)
  make.unique(x, sep = "_")
}

save_excel_bundle <- function(sheet_list, manifest_df, out_xlsx) {
  sheet_list <- Filter(Negate(is.null), sheet_list)
  names(sheet_list) <- safe_sheet_names(names(sheet_list))
  
  wb <- createWorkbook()
  addWorksheet(wb, "manifest")
  writeData(wb, "manifest", manifest_df)
  
  for (nm in names(sheet_list)) {
    addWorksheet(wb, nm)
    writeData(wb, nm, sheet_list[[nm]])
  }
  
  saveWorkbook(wb, out_xlsx, overwrite = TRUE)
  invisible(out_xlsx)
}

theme_marine <- function(base_size = BASE) {
  theme_bw(base_size = base_size) +
    theme(
      text = element_text(size = base_size, colour = "black"),
      axis.text.x = element_text(size = base_size, colour = "black"),
      axis.text.y = element_text(size = base_size, colour = "black"),
      axis.title.x = element_text(size = base_size, colour = "black"),
      axis.title.y = element_text(size = base_size, colour = "black"),
      legend.text = element_text(size = base_size, colour = "black"),
      legend.title = element_text(size = base_size, colour = "black"),
      strip.text = element_text(size = base_size, colour = "black"),
      strip.background = element_blank(),
      strip.placement = "outside",
      panel.grid = element_blank(),
      panel.grid.major.x = element_line(linewidth = .1, colour = "grey85", linetype = 2),
      panel.grid.major.y = element_line(linewidth = .1, colour = "grey85", linetype = 2)
    )
}

distribution_ci_default <- function() {
  2.7 * 1000 / 5800
}

extract_legend_grob <- function(p) {
  gt <- ggplotGrob(p)
  idx <- which(sapply(gt$grobs, function(x) x$name) == "guide-box")
  if (length(idx) == 0) return(nullGrob())
  gt$grobs[[idx[1]]]
}

norm_key <- function(x) {
  x %>%
    as.character() %>%
    str_to_lower() %>%
    str_replace_all("[^a-z0-9]+", " ") %>%
    str_squish()
}

###############################################################################
# Fig. 3 - HSFO
# Panel A: refinery-level domestic HSFO WTT CI
# Panel B: voyage-level exported HSFO WTT CI
###############################################################################

source("marine_figures_shared_helpers.R")

# ---------------------------------------------------------------------------
# User paths
# ---------------------------------------------------------------------------
fig_dir <- "PATH/Fig 3"
dir.create(fig_dir, recursive = TRUE, showWarnings = FALSE)

inp_ci_d <- "PATH/16_hsfo_CI_upstream_domestic.csv"
inp_ci0  <- "PATH/12_hsfo_CI_upstream_shipping_1_event.csv"
inp_archie <- "PATH/Archie_v3_Backend Generator_final 2023.xlsx"

stopifnot(file.exists(inp_ci_d), file.exists(inp_ci0), file.exists(inp_archie))

# ---------------------------------------------------------------------------
# Load inputs
# ---------------------------------------------------------------------------
ci_d <- readr::read_csv(inp_ci_d, show_col_types = FALSE) %>% as.data.frame()
ci0  <- readr::read_csv(inp_ci0,  show_col_types = FALSE) %>% as.data.frame()

archie_ref <- readxl::read_excel(inp_archie, sheet = "out_refineries") %>%
  transmute(
    refinery_id     = str_squish(`Refinery ID`),
    archie_ref_name = str_squish(`Short Refinery Name`),
    archie_country  = str_squish(`Country`)
  )

# ---------------------------------------------------------------------------
# Panel A
# ---------------------------------------------------------------------------
distribution_CI_gCO2e_MJ <- distribution_ci_default()

ci_d2 <- ci_d %>%
  mutate(refinery_id = str_squish(refinery_id)) %>%
  left_join(archie_ref, by = "refinery_id") %>%
  mutate(
    archie_ref_name = na_if(str_squish(archie_ref_name), ""),
    archie_country  = na_if(str_squish(archie_country), ""),
    archie_ref_name = coalesce(archie_ref_name, str_squish(prelim_refinery_name)),
    archie_country  = coalesce(archie_country, str_squish(Country_2))
  )

df_A <- ci_d2 %>%
  transmute(
    refinery_name = coalesce(na_if(str_squish(archie_ref_name), ""), as.character(refinery_id)),
    Region,
    Country_1,
    w_dom_bpd = nz_num(Domestic_rem_WM_hsfo_Vol_bpd),
    upstream  = nz_num(crude_intake_upstream_ci_g_co2e_mj),
    transport = nz_num(crude_transport_ci_g_co2e_mj),
    refining  = nz_num(hsfo_refining_ci_g_co2e_mj)
  ) %>%
  filter(!is.na(refinery_name), nzchar(refinery_name), w_dom_bpd > 0)

global_dom <- df_A %>%
  summarise(
    V_dom_bpd         = sum(w_dom_bpd, na.rm = TRUE),
    CI_upstream_wavg  = safe_div(sum(upstream  * w_dom_bpd, na.rm = TRUE), V_dom_bpd),
    CI_transport_wavg = safe_div(sum(transport * w_dom_bpd, na.rm = TRUE), V_dom_bpd),
    CI_refining_wavg  = safe_div(sum(refining  * w_dom_bpd, na.rm = TRUE), V_dom_bpd)
  ) %>%
  mutate(
    CI_distribution  = distribution_CI_gCO2e_MJ,
    CI_total_bs_wavg = CI_upstream_wavg + CI_transport_wavg + CI_refining_wavg + CI_distribution
  )

global_avg_CI <- as.numeric(global_dom$CI_total_bs_wavg)

ref_df <- df_A %>%
  group_by(refinery = refinery_name) %>%
  summarise(
    loaded_dom_bpd    = sum(w_dom_bpd, na.rm = TRUE),
    upstream_CI_wavg  = safe_div(sum(upstream  * w_dom_bpd, na.rm = TRUE), loaded_dom_bpd),
    transport_CI_wavg = safe_div(sum(transport * w_dom_bpd, na.rm = TRUE), loaded_dom_bpd),
    refining_CI_wavg  = safe_div(sum(refining  * w_dom_bpd, na.rm = TRUE), loaded_dom_bpd),
    .groups = "drop"
  ) %>%
  mutate(
    distribution_CI   = distribution_CI_gCO2e_MJ,
    total_CI_wavg     = upstream_CI_wavg + transport_CI_wavg + refining_CI_wavg + distribution_CI
  ) %>%
  arrange(total_CI_wavg) %>%
  mutate(
    produced_volume_Mt = loaded_dom_bpd * 0.0001364,
    cumulative_prod_Mt = cumsum(produced_volume_Mt)
  )

plot_df_A <- ref_df %>%
  mutate(
    offset_0 = 0,
    offset_1 = upstream_CI_wavg,
    offset_2 = offset_1 + transport_CI_wavg,
    offset_3 = offset_2 + refining_CI_wavg,
    offset_4 = offset_3 + distribution_CI
  )

ord_A <- c("upstream", "refining", "transport", "distribution")
stage_cols_A <- setNames(ggsci::pal_aaas("default")(length(ord_A)), ord_A)

global_stage_vals_A <- tibble(
  stage = factor(ord_A, levels = ord_A),
  value = c(
    global_dom$CI_upstream_wavg,
    global_dom$CI_refining_wavg,
    global_dom$CI_transport_wavg,
    distribution_CI_gCO2e_MJ
  )
) %>%
  mutate(share = value / sum(value))

p_bar_A <- ggplot(global_stage_vals_A, aes(stage, value, fill = stage)) +
  geom_col(width = 0.7, colour = "black", linewidth = 0.2) +
  scale_fill_manual(values = stage_cols_A, breaks = ord_A, name = "") +
  scale_x_discrete(limits = ord_A) +
  labs(
    x = "",
    y = expression(atop("Global VWA WTT CI", "(domestic) gCO"[2]~"e MJ"^-1*"")),
    title = "a. Well-to-Tank stages CI"
  ) +
  theme_bw(base_size = BASE) +
  theme(
    axis.text.x = element_blank(),
    axis.ticks.x = element_blank(),
    legend.position = "none",
    panel.grid.major = element_blank(),
    panel.grid.minor = element_blank(),
    plot.title = element_text(size = BASE)
  )

bar_grob_A <- ggplotGrob(p_bar_A)

p_donut_A <- ggplot(global_stage_vals_A, aes(x = 2, y = share, fill = stage)) +
  geom_col(width = 1, colour = "black", linewidth = 0.2) +
  coord_polar(theta = "y") +
  xlim(1, 2.8) +
  scale_fill_manual(values = stage_cols_A, breaks = ord_A, name = "") +
  theme_void(base_size = BASE) +
  labs(title = "b. Domestic VWA WTT CI share (%)") +
  geom_text(
    aes(label = percent(share, accuracy = 1)),
    position = position_stack(vjust = 0.5),
    size = pt_size(BASE * 0.9),
    colour = "black"
  ) +
  theme(
    legend.position = "none",
    plot.title = element_text(size = BASE, hjust = 0)
  )

donut_grob_A <- ggplotGrob(p_donut_A)

top_labels_A <- plot_df_A %>%
  slice_max(order_by = produced_volume_Mt, n = 15, with_ties = FALSE) %>%
  pull(refinery) %>%
  unique()

top_labels_A <- unique(c(top_labels_A, "Bahrain"))

p_ref <- ggplot(plot_df_A) +
  geom_rect(aes(xmin = cumulative_prod_Mt - produced_volume_Mt, xmax = cumulative_prod_Mt,
                ymin = offset_0, ymax = offset_1, fill = "upstream"),
            colour = "black", linewidth = 0.01) +
  geom_rect(aes(xmin = cumulative_prod_Mt - produced_volume_Mt, xmax = cumulative_prod_Mt,
                ymin = offset_1, ymax = offset_2, fill = "transport"),
            colour = "black", linewidth = 0.01) +
  geom_rect(aes(xmin = cumulative_prod_Mt - produced_volume_Mt, xmax = cumulative_prod_Mt,
                ymin = offset_2, ymax = offset_3, fill = "refining"),
            colour = "black", linewidth = 0.01) +
  geom_rect(aes(xmin = cumulative_prod_Mt - produced_volume_Mt, xmax = cumulative_prod_Mt,
                ymin = offset_3, ymax = offset_4, fill = "distribution"),
            colour = "black", linewidth = 0.01) +
  scale_x_continuous(
    limits = c(0, max(plot_df_A$cumulative_prod_Mt, na.rm = TRUE)),
    breaks = c(0, 0.25, 0.5, 0.75, 1) * max(plot_df_A$cumulative_prod_Mt, na.rm = TRUE),
    labels = c(0, 25, 50, 75, 100),
    name = "Cumulative HSFO produced for domestic market (%, ~1.6 Mbpd, ~88 Mt y⁻¹)\nHSFO density assumption (bbl/t ≈ 6.6)",
    expand = c(0.01, 0.01)
  ) +
  scale_y_continuous(
    expand = c(0.02, 0.05),
    limits = c(0, max(plot_df_A$offset_4, na.rm = TRUE) * 1.1),
    name = expression(atop("Domestic refinery-level HSFO Well-to-Tank", "Carbon Intensity (gCO"[2]~"e MJ"^-1*")"))
  ) +
  scale_fill_manual(values = stage_cols_A, breaks = ord_A, name = "") +
  labs(title = "A. Refinery-level HSFO CI for domestic use") +
  theme_marine(BASE) +
  theme(
    plot.title = element_text(size = BASE, hjust = 0),
    legend.position = c(0.02, 0.98),
    legend.justification = c(0, 1),
    legend.direction = "horizontal",
    legend.key.width = unit(1.2, "cm")
  ) +
  annotation_custom(
    bar_grob_A,
    xmin = max(plot_df_A$cumulative_prod_Mt, na.rm = TRUE) * 0.05,
    xmax = max(plot_df_A$cumulative_prod_Mt, na.rm = TRUE) * 0.45,
    ymin = max(plot_df_A$offset_4, na.rm = TRUE) * 0.30,
    ymax = max(plot_df_A$offset_4, na.rm = TRUE) * 0.96
  ) +
  annotation_custom(
    donut_grob_A,
    xmin = max(plot_df_A$cumulative_prod_Mt, na.rm = TRUE) * 0.50,
    xmax = max(plot_df_A$cumulative_prod_Mt, na.rm = TRUE) * 0.90,
    ymin = max(plot_df_A$offset_4, na.rm = TRUE) * 0.30,
    ymax = max(plot_df_A$offset_4, na.rm = TRUE) * 0.98
  ) +
  geom_hline(yintercept = global_avg_CI, linetype = "dashed", colour = "blue", linewidth = 0.8) +
  annotate(
    "text",
    x = max(plot_df_A$cumulative_prod_Mt, na.rm = TRUE) * 0.15,
    y = global_avg_CI + 7,
    label = paste0("Global domestic VWA CI (before shipping) = ", sprintf("%.1f", global_avg_CI), " gCO₂e/MJ"),
    colour = "black",
    size = pt_size(BASE * 0.9),
    hjust = 0
  ) +
  geom_text(
    data = plot_df_A %>% filter(refinery %in% top_labels_A),
    aes(x = cumulative_prod_Mt - produced_volume_Mt / 2, y = 0, label = refinery),
    angle = 90,
    size = pt_size(BASE * 0.9),
    hjust = -0.02,
    vjust = 0.5,
    colour = "black"
  )

# ---------------------------------------------------------------------------
# Panel B
# ---------------------------------------------------------------------------
ci_B <- ci0 %>%
  mutate(
    refinery = if_else(
      is.na(load_port_refineries_list_1) | load_port_refineries_list_1 == "",
      NA_character_,
      str_trim(sub("\\|.*$", "", load_port_refineries_list_1))
    ),
    country_load = loading_country_territory_1,
    country_discharge = discharging_country_territory_1,
    country_load_ISO3 = countrycode(country_load, "country.name", "iso3c", warn = TRUE),
    country_discharge_ISO3 = countrycode(country_discharge, "country.name", "iso3c", warn = TRUE)
  ) %>%
  mutate(
    refinery_clean = str_squish(coalesce(refinery, "")),
    load_iso_clean = str_squish(coalesce(country_load_ISO3, "")),
    dport_clean    = str_squish(coalesce(discharging_port_1, "")),
    diso_clean     = str_squish(coalesce(country_discharge_ISO3, "")),
    route = if_else(
      refinery_clean != "" & load_iso_clean != "" & dport_clean != "" & diso_clean != "",
      paste0(refinery_clean, " (", load_iso_clean, ", ref) → ", dport_clean, " (", diso_clean, ", port)"),
      NA_character_
    )
  ) %>%
  select(-refinery_clean, -load_iso_clean, -dport_clean, -diso_clean)

voyage_wtt <- ci_B %>%
  mutate(
    loaded_t  = nz_num(loading_vol_tons_1),
    upstream  = nz_num(load_port_upstream_ci_g_per_MJ_1),
    transport = nz_num(load_port_crude_transport_ci_g_MJ_1),
    refining  = nz_num(load_port_refining_ci_g_per_MJ_1),
    shipping  = nz_num(ci_shiping_total_gCO2_per_MJ)
  ) %>%
  filter(!is.na(voyage_id), loaded_t > 0) %>%
  group_by(voyage_id) %>%
  summarise(
    produced_volume_t = sum(loaded_t, na.rm = TRUE),
    upstream_CI_wavg  = weighted.mean(upstream,  loaded_t, na.rm = TRUE),
    transport_CI_wavg = weighted.mean(transport, loaded_t, na.rm = TRUE),
    refining_CI_wavg  = weighted.mean(refining,  loaded_t, na.rm = TRUE),
    shipping_CI_wavg  = weighted.mean(shipping,  loaded_t, na.rm = TRUE),
    route_primary = {
      x <- tibble(r = route, w = loaded_t) %>%
        filter(!is.na(r) & r != "") %>%
        group_by(r) %>%
        summarise(w = sum(w, na.rm = TRUE), .groups = "drop") %>%
        arrange(desc(w)) %>%
        slice(1) %>%
        pull(r)
      if (length(x) == 0) NA_character_ else x
    },
    route_all = {
      u <- unique(route[!is.na(route) & route != ""])
      if (length(u) == 0) NA_character_ else paste(u, collapse = " | ")
    },
    .groups = "drop"
  ) %>%
  mutate(
    distribution_CI_gCO2e_MJ = distribution_CI_gCO2e_MJ,
    total_CI_wavg     = upstream_CI_wavg + transport_CI_wavg + refining_CI_wavg + shipping_CI_wavg,
    total_WTT_CI_wavg = total_CI_wavg + distribution_CI_gCO2e_MJ
  ) %>%
  as.data.frame()

global_stage_B <- voyage_wtt %>%
  summarise(
    upstream     = weighted.mean(upstream_CI_wavg,        produced_volume_t, na.rm = TRUE),
    transport    = weighted.mean(transport_CI_wavg,       produced_volume_t, na.rm = TRUE),
    refining     = weighted.mean(refining_CI_wavg,        produced_volume_t, na.rm = TRUE),
    shipping     = weighted.mean(shipping_CI_wavg,        produced_volume_t, na.rm = TRUE),
    distribution = weighted.mean(distribution_CI_gCO2e_MJ,produced_volume_t, na.rm = TRUE)
  ) %>%
  pivot_longer(everything(), names_to = "Stage", values_to = "Global_CI") %>%
  mutate(Global_CI = round(Global_CI, 1))

global_avg_CI_B <- round(sum(global_stage_B$Global_CI), 1)

plot_df_B <- voyage_wtt %>%
  arrange(total_CI_wavg) %>%
  mutate(
    produced_volume_Mt = produced_volume_t / 1e6,
    cumulative_prod_Mt = cumsum(produced_volume_Mt),
    offset_0 = 0,
    offset_1 = upstream_CI_wavg,
    offset_2 = offset_1 + transport_CI_wavg,
    offset_3 = offset_2 + refining_CI_wavg,
    offset_4 = offset_3 + shipping_CI_wavg,
    offset_5 = offset_4 + distribution_CI_gCO2e_MJ
  )

p_voy1 <- ggplot(plot_df_B) +
  geom_rect(aes(xmin = cumulative_prod_Mt - produced_volume_Mt, xmax = cumulative_prod_Mt,
                ymin = offset_0, ymax = offset_1, fill = "upstream"),
            colour = NA) +
  geom_rect(aes(xmin = cumulative_prod_Mt - produced_volume_Mt, xmax = cumulative_prod_Mt,
                ymin = offset_1, ymax = offset_2, fill = "transport"),
            colour = NA) +
  geom_rect(aes(xmin = cumulative_prod_Mt - produced_volume_Mt, xmax = cumulative_prod_Mt,
                ymin = offset_2, ymax = offset_3, fill = "refining"),
            colour = NA) +
  geom_rect(aes(xmin = cumulative_prod_Mt - produced_volume_Mt, xmax = cumulative_prod_Mt,
                ymin = offset_3, ymax = offset_4, fill = "shipping"),
            colour = NA) +
  geom_rect(aes(xmin = cumulative_prod_Mt - produced_volume_Mt, xmax = cumulative_prod_Mt,
                ymin = offset_4, ymax = offset_5, fill = "distribution"),
            colour = NA) +
  scale_y_continuous(
    name = expression(atop("Voyage-level Well-to-Tank", "Carbon Intensity (gCO"[2]~"e MJ"^-1*")")),
    limits = c(0, max(plot_df_B$offset_5, na.rm = TRUE) * 1.1),
    expand = c(0.02, 0.02)
  ) +
  scale_x_continuous(
    limits = c(0, max(plot_df_B$cumulative_prod_Mt, na.rm = TRUE)),
    breaks = c(0, 0.25, 0.5, 0.75, 1) * max(plot_df_B$cumulative_prod_Mt, na.rm = TRUE),
    labels = c(0, 25, 50, 75, 100),
    name = "Cumulative HSFO exported for international market (%, ~4.9 Mbpd, ~270 Mt y⁻¹)\nHSFO density assumption (bbl/t ≈ 6.6)",
    expand = c(0.01, 0.01)
  ) +
  geom_hline(yintercept = global_avg_CI_B, linetype = "dashed", colour = "blue", linewidth = 0.8) +
  annotate(
    "text",
    x = max(plot_df_B$cumulative_prod_Mt, na.rm = TRUE) * 0.03,
    y = global_avg_CI_B + 1,
    label = paste0("Global shipped VWA WTT CI = ", sprintf("%.1f", global_avg_CI_B), " gCO₂e/MJ"),
    colour = "black",
    size = pt_size(BASE * 0.9),
    hjust = 0
  ) +
  scale_fill_npg(name = "", breaks = c("upstream", "transport", "refining", "shipping", "distribution")) +
  theme_marine(BASE) +
  theme(
    plot.title = element_text(hjust = 0, size = BASE),
    legend.position = c(0.02, 0.98),
    legend.justification = c(0, 1),
    legend.direction = "horizontal",
    legend.key.width = unit(1.2, "cm")
  ) +
  labs(title = "B. Voyage-level HSFO WTT CI for international trade")

df_violin <- ci0 %>%
  transmute(
    exporter = str_trim(loading_country_territory_1),
    vol_t    = nz_num(loading_vol_tons_1),
    total_ci = nz_num(ci_shiping_total_gCO2_per_MJ) + nz_num(load_port_total_before_ship_g_MJ_1)
  ) %>%
  filter(!is.na(exporter), vol_t > 0, is.finite(total_ci))

top_exporters <- df_violin %>%
  group_by(exporter) %>%
  summarise(load_t = sum(vol_t, na.rm = TRUE), .groups = "drop") %>%
  slice_max(load_t, n = 5, with_ties = FALSE) %>%
  pull(exporter)

map_iso3 <- function(x) {
  iso <- countrycode(x, origin = "country.name", destination = "iso3c", warn = TRUE)
  ifelse(is.na(iso) | iso == "", x, iso)
}

df2_violin <- df_violin %>%
  filter(exporter %in% top_exporters) %>%
  mutate(iso3 = map_iso3(exporter)) %>%
  mutate(iso3 = if_else(iso3 == "ARE", "UAE", iso3))

x_order <- df2_violin %>%
  group_by(iso3) %>%
  summarise(load_t = sum(vol_t, na.rm = TRUE), .groups = "drop") %>%
  arrange(desc(load_t)) %>%
  pull(iso3)

global_exp_totals <- ci0 %>%
  transmute(
    exporter = str_trim(loading_country_territory_1),
    cargo_t  = nz_num(total_cargo_t)
  ) %>%
  filter(!is.na(exporter), cargo_t > 0) %>%
  group_by(exporter) %>%
  summarise(cargo_t = sum(cargo_t, na.rm = TRUE), .groups = "drop")

global_total_cargo <- sum(global_exp_totals$cargo_t, na.rm = TRUE)

share_tbl <- global_exp_totals %>%
  filter(exporter %in% top_exporters) %>%
  mutate(iso3 = map_iso3(exporter)) %>%
  mutate(iso3 = if_else(iso3 == "ARE", "UAE", iso3)) %>%
  group_by(iso3) %>%
  summarise(cargo_t = sum(cargo_t, na.rm = TRUE), .groups = "drop") %>%
  mutate(share = if (global_total_cargo > 0) cargo_t / global_total_cargo else 0) %>%
  arrange(desc(share))

col_vec <- c(
  RUS = "#8AA1C1",
  MYS = "#F3A7A3",
  UAE = "#7FBF9E",
  SAU = "#B69CCB",
  SGP = "#84C2BE",
  MEX = "#F0B08C",
  NLD = "#A3A0C9",
  BRA = "#C4C87C",
  KWT = "#E9B77C",
  TUR = "#9BB9E2"
)
col_vec <- col_vec[x_order]

lab_vec <- setNames(
  paste0(x_order, " (", percent(share_tbl$share[match(x_order, share_tbl$iso3)], accuracy = 1), ")"),
  x_order
)

p_violin <- ggplot(
  df2_violin,
  aes(x = factor(iso3, levels = x_order), y = total_ci, weight = vol_t, fill = iso3)
) +
  geom_violin(scale = "width", trim = TRUE, colour = "grey") +
  scale_fill_manual(values = col_vec, breaks = x_order, labels = lab_vec, name = "Export share (top-5)") +
  labs(
    title = "a. CI variation of HSFO cargoes by top exporting countries",
    x = NULL,
    y = expression("Well-to-Tank CI (gCO"[2]*"e MJ"^-1*")")
  ) +
  theme_bw(base_size = BASE) +
  theme(
    panel.grid.minor = element_blank(),
    axis.text.x = element_text(size = BASE),
    axis.text.y = element_text(size = BASE),
    plot.title = element_text(size = BASE, hjust = 0),
    legend.position = "right"
  )

violin_grob <- ggplotGrob(p_violin)

p_voy_v <- p_voy1 +
  annotation_custom(
    violin_grob,
    xmin = 0.05 * max(plot_df_B$cumulative_prod_Mt, na.rm = TRUE),
    xmax = 0.60 * max(plot_df_B$cumulative_prod_Mt, na.rm = TRUE),
    ymin = 0.40 * max(plot_df_B$offset_5, na.rm = TRUE),
    ymax = 0.98 * max(plot_df_B$offset_5, na.rm = TRUE)
  )

# ---------------------------------------------------------------------------
# Example route labels
# ---------------------------------------------------------------------------
targets <- c(3.716526, 6.001596, 7.5, 10.01282, 12, 12.8, 14, 15.18522, 17.18222, 20.30496, 32.68776, 38.95414)
tol <- 0.001

top_routes <- plot_df_B %>%
  filter(Reduce(`|`, lapply(targets, function(v) near(total_WTT_CI_wavg, v, tol = tol)))) %>%
  arrange(desc(total_WTT_CI_wavg)) %>%
  filter(!route_primary %in% c(
    "Madero (MEX, ref) → Topolobampo (MEX, port)",
    "Antwerp (ExxonMobil) (BEL, ref) → Rotterdam (NLD, port)",
    "Mozyr (POL, ref) → Rotterdam (NLD, port)",
    "Kirishi (RUS, ref) → Kalamata (GRC, port)",
    "Gonfreville l'Orcher (FRA, ref) → Mongstad (NOR, port)",
    "Rotterdam (Vitol) (NLD, ref) → LOOP Terminal (USA, port)"
  )) %>%
  filter(!voyage_id %in% c(
    "L-9798868-2023523-105316-202352-131713",
    "L-9654919-2023430-5415-2023411-25842"
  )) %>%
  mutate(
    route_all = ifelse(route_all == "Fujairah (ARE, ref) → Yantai (CHN, port)", "Fujairah (UAE, ref) → Yantai (CHN, port)", route_all),
    route_all = ifelse(route_all == "Ras Tanura (SAU, ref) → Fujairah (ARE, port)", "Ras Tanura (SAU, ref) → Fujairah (UAE, port)", route_all),
    route_primary = ifelse(route_primary == "Fujairah (ARE, ref) → Yantai (CHN, port)", "Fujairah (UAE, ref) → Yantai (CHN, port)", route_primary),
    route_primary = ifelse(route_primary == "Ras Tanura (SAU, ref) → Fujairah (ARE, port)", "Ras Tanura (SAU, ref) → Fujairah (UAE, port)", route_primary)
  )

ycol <- if ("total_CI_wavg" %in% names(top_routes)) "total_CI_wavg" else "total_WTT_CI_wavg"

top_routes2 <- top_routes %>%
  mutate(
    x_lab2 = cumulative_prod_Mt,
    x_lab = cumulative_prod_Mt,
    y_lab = .data[[ycol]] + 3,
    y_lab2 = .data[[ycol]],
    y_lab3 = .data[[ycol]] + 1.5
  )

p_voy_v2 <- p_voy_v +
  geom_text(
    data = top_routes2,
    aes(x = x_lab, y = y_lab, label = route_primary),
    size = pt_size(BASE * 0.9),
    hjust = 0.65,
    vjust = 1.7,
    colour = "black",
    inherit.aes = FALSE
  ) +
  geom_segment(
    data = top_routes2,
    aes(x = x_lab2, xend = x_lab2, y = y_lab2, yend = y_lab3),
    arrow = arrow(length = unit(0.2, "cm"), type = "closed"),
    colour = "black",
    inherit.aes = FALSE
  )

# ---------------------------------------------------------------------------
# Combine and save
# ---------------------------------------------------------------------------
combined <- (p_ref + theme(plot.margin = margin(8, 12, 6, 8))) /
  (p_voy_v2 + theme(plot.margin = margin(6, 12, 8, 8))) +
  plot_layout(heights = c(1, 1), guides = "keep")

ggsave(file.path(fig_dir, "Fig 3.png"), combined, width = 12, height = 13, dpi = 300)
ggsave(file.path(fig_dir, "Fig 3.svg"), combined, width = 12, height = 13, dpi = 300)

# ---------------------------------------------------------------------------
# Excel bundle
# ---------------------------------------------------------------------------
inputs_manifest <- tibble(
  kind = "input_file",
  name = c("ci_d_domestic", "ci0_shipping", "archie_refineries_xlsx"),
  path = c(inp_ci_d, inp_ci0, inp_archie),
  exists = file.exists(c(inp_ci_d, inp_ci0, inp_archie))
)

outputs_manifest <- bind_rows(
  tibble(kind = "output_df", name = "A_ci_d2_joined", path = NA_character_, n_rows = nrow(ci_d2), n_cols = ncol(ci_d2)),
  tibble(kind = "output_df", name = "A_domestic_working_df", path = NA_character_, n_rows = nrow(df_A), n_cols = ncol(df_A)),
  tibble(kind = "output_df", name = "A_refinery_level_wavgs", path = NA_character_, n_rows = nrow(ref_df), n_cols = ncol(ref_df)),
  tibble(kind = "output_df", name = "A_plot_df_stacked", path = NA_character_, n_rows = nrow(plot_df_A), n_cols = ncol(plot_df_A)),
  tibble(kind = "output_df", name = "A_global_dom_wavgs", path = NA_character_, n_rows = nrow(global_dom), n_cols = ncol(global_dom)),
  tibble(kind = "output_df", name = "B_raw_voyage_events", path = NA_character_, n_rows = nrow(ci0), n_cols = ncol(ci0)),
  tibble(kind = "output_df", name = "B_enriched_routes_iso3", path = NA_character_, n_rows = nrow(ci_B), n_cols = ncol(ci_B)),
  tibble(kind = "output_df", name = "B_voyage_wtt_wavgs", path = NA_character_, n_rows = nrow(voyage_wtt), n_cols = ncol(voyage_wtt)),
  tibble(kind = "output_df", name = "B_global_stage_means", path = NA_character_, n_rows = nrow(global_stage_B), n_cols = ncol(global_stage_B)),
  tibble(kind = "output_df", name = "B_plot_df_stacked", path = NA_character_, n_rows = nrow(plot_df_B), n_cols = ncol(plot_df_B)),
  tibble(kind = "output_df", name = "B_top_exporters_violin_df", path = NA_character_, n_rows = nrow(df2_violin), n_cols = ncol(df2_violin)),
  tibble(kind = "output_df", name = "B_top_exporters_global_shares", path = NA_character_, n_rows = nrow(share_tbl), n_cols = ncol(share_tbl)),
  tibble(kind = "output_df", name = "B_labelled_example_routes", path = NA_character_, n_rows = nrow(top_routes2), n_cols = ncol(top_routes2))
)

figs_manifest <- tibble(
  kind = "figure_file",
  name = c("Fig 3 PNG", "Fig 3 SVG"),
  path = file.path(fig_dir, c("Fig 3.png", "Fig 3.svg")),
  exists = file.exists(file.path(fig_dir, c("Fig 3.png", "Fig 3.svg")))
)

fig3_manifest <- bind_rows(inputs_manifest, outputs_manifest, figs_manifest)

sheet_list <- list(
  A_ci_d2_joined = ci_d2,
  A_domestic_working_df = df_A,
  A_refinery_level_wavgs = ref_df,
  A_plot_df_stacked = plot_df_A,
  A_global_dom_wavgs = global_dom,
  B_raw_voyage_events = ci0,
  B_enriched_routes_iso3 = ci_B,
  B_voyage_wtt_wavgs = voyage_wtt,
  B_global_stage_means = global_stage_B,
  B_plot_df_stacked = plot_df_B,
  B_top_exporters_violin_df = df2_violin,
  B_top_exporters_global_shares = share_tbl,
  B_labelled_example_routes = top_routes2
)

save_excel_bundle(
  sheet_list = sheet_list,
  manifest_df = fig3_manifest,
  out_xlsx = file.path(fig_dir, "Fig3_data_bundle.xlsx")
)

###############################################################################
# Fig. 4 - LPG
# Panel A: refinery-level LPG CI before shipping
# Panel B: asset-level variability across top producers
###############################################################################

source("marine_figures_shared_helpers.R")

# ---------------------------------------------------------------------------
# User paths
# ---------------------------------------------------------------------------
fig_dir <- "PATH/Fig 4"
dir.create(fig_dir, recursive = TRUE, showWarnings = FALSE)

file_path <- "PATH/Marine fuel well to refinery exit CI v5 (PRELIMv20) LPG lower v2 2023.xlsx"
inp_archie <- "PATH/Archie_v3_Backend Generator_final 2023.xlsx"

stopifnot(file.exists(file_path), file.exists(inp_archie))

# ---------------------------------------------------------------------------
# Load and clean
# ---------------------------------------------------------------------------
df_ci <- readxl::read_excel(file_path, sheet = 1) %>%
  janitor::clean_names() %>%
  as.data.frame()

df_ci <- df_ci %>%
  rename(
    refinery_ID = x1,
    energy_density_LHV_mj_bbl = mj_bbl
  ) %>%
  select(-any_of(c(
    "x8", "x11", "x18",
    "hsfo_domestic_use_bbl_d",
    "hsfo_import_bbl_d",
    "fuel_trade_shipping_ci_g_co2e_mj",
    "combustion_ci_g_co2e_mj"
  )))

lpg <- df_ci %>%
  select(-any_of(c("hsfo_volume_bbl_d", "hsfo_refining_ci_g_co2e_mj", "hsfo_total_before_shipping")))

lpg_clean <- lpg %>%
  filter(
    !is.na(lpg_volume_bbl_d),
    !is.na(lpg_refining_ci_g_co2e_mj),
    !is.na(lpg_total_before_shipping)
  )

archie_ref <- readxl::read_excel(inp_archie, sheet = "out_refineries") %>%
  transmute(
    refinery_ID     = str_squish(`Refinery ID`),
    archie_ref_name = str_squish(`Short Refinery Name`),
    archie_country  = str_squish(Country)
  )

lpg_clean_2 <- lpg_clean %>%
  mutate(refinery_ID = str_squish(refinery_ID)) %>%
  left_join(archie_ref, by = "refinery_ID")

readr::write_csv(lpg_clean_2, file.path(fig_dir, "01_lpg_well_to_refinery_CI_asset_level.csv"))

# ---------------------------------------------------------------------------
# Panel A
# ---------------------------------------------------------------------------
distribution_CI_gCO2e_MJ <- distribution_ci_default()

lpg_df <- lpg_clean_2 %>%
  transmute(
    refinery_name = coalesce(archie_ref_name, as.character(refinery_ID)),
    country       = country,
    w_dom_bbl_d   = nz_num(lpg_volume_bbl_d),
    upstream      = nz_num(crude_intake_upstream_ci_g_co2e_mj),
    transport     = nz_num(crude_transport_ci_g_co2e_mj),
    refining      = nz_num(lpg_refining_ci_g_co2e_mj),
    LHV_MJ_per_bbl = nz_num(energy_density_LHV_mj_bbl)
  ) %>%
  filter(!is.na(refinery_name), nzchar(refinery_name), w_dom_bbl_d > 0)

global_dom_lpg <- lpg_df %>%
  summarise(
    V_dom_bbl_d       = sum(w_dom_bbl_d, na.rm = TRUE),
    CI_upstream_wavg  = safe_div(sum(upstream  * w_dom_bbl_d, na.rm = TRUE), V_dom_bbl_d),
    CI_transport_wavg = safe_div(sum(transport * w_dom_bbl_d, na.rm = TRUE), V_dom_bbl_d),
    CI_refining_wavg  = safe_div(sum(refining  * w_dom_bbl_d, na.rm = TRUE), V_dom_bbl_d)
  ) %>%
  mutate(
    CI_distribution  = distribution_CI_gCO2e_MJ,
    CI_total_bs_wavg = CI_upstream_wavg + CI_transport_wavg + CI_refining_wavg + CI_distribution
  )

global_avg_CI_lpg <- as.numeric(global_dom_lpg$CI_total_bs_wavg)

ref_df_lpg <- lpg_df %>%
  group_by(refinery = refinery_name) %>%
  summarise(
    loaded_dom_bbl_d  = sum(w_dom_bbl_d, na.rm = TRUE),
    upstream_CI_wavg  = safe_div(sum(upstream  * w_dom_bbl_d, na.rm = TRUE), loaded_dom_bbl_d),
    transport_CI_wavg = safe_div(sum(transport * w_dom_bbl_d, na.rm = TRUE), loaded_dom_bbl_d),
    refining_CI_wavg  = safe_div(sum(refining  * w_dom_bbl_d, na.rm = TRUE), loaded_dom_bbl_d),
    .groups = "drop"
  ) %>%
  mutate(
    distribution_CI = distribution_CI_gCO2e_MJ,
    total_CI_wavg   = upstream_CI_wavg + transport_CI_wavg + refining_CI_wavg + distribution_CI
  ) %>%
  arrange(total_CI_wavg) %>%
  mutate(
    produced_volume_Mt = loaded_dom_bbl_d * 0.0001364,
    cumulative_prod_Mt = cumsum(produced_volume_Mt)
  )

plot_df_lpg <- ref_df_lpg %>%
  mutate(
    offset_0 = 0,
    offset_1 = upstream_CI_wavg,
    offset_2 = offset_1 + transport_CI_wavg,
    offset_3 = offset_2 + refining_CI_wavg,
    offset_4 = offset_3 + distribution_CI
  )

ord_lpg <- c("refining", "upstream", "transport", "distribution")
stage_cols_lpg <- setNames(ggsci::pal_bmj("default")(length(ord_lpg)), ord_lpg)

global_stage_vals_lpg <- tibble(
  stage = factor(ord_lpg, levels = ord_lpg),
  value = c(
    global_dom_lpg$CI_refining_wavg,
    global_dom_lpg$CI_upstream_wavg,
    global_dom_lpg$CI_transport_wavg,
    distribution_CI_gCO2e_MJ
  )
) %>%
  mutate(share = value / sum(value))

p_bar_lpg <- ggplot(global_stage_vals_lpg, aes(stage, value, fill = stage)) +
  geom_col(width = 0.7, colour = "black", linewidth = 0.2) +
  scale_fill_manual(values = stage_cols_lpg, breaks = ord_lpg, name = "") +
  scale_x_discrete(limits = ord_lpg) +
  labs(
    x = "",
    y = expression(atop("Global VWA WTT CI", "(refinery-sourced LPG) gCO"[2]~"e MJ"^-1*"")),
    title = "a. Well-to-Tank stages CI"
  ) +
  theme_bw(base_size = BASE) +
  theme(
    axis.text.x = element_blank(),
    axis.ticks.x = element_blank(),
    legend.position = "none",
    plot.title = element_text(size = BASE),
    panel.grid.major = element_blank(),
    panel.grid.minor = element_blank()
  )

bar_grob_lpg <- ggplotGrob(p_bar_lpg)

p_donut_lpg <- ggplot(global_stage_vals_lpg, aes(x = 2, y = share, fill = stage)) +
  geom_col(width = 1, colour = "black", linewidth = 0.2) +
  coord_polar(theta = "y") +
  xlim(1, 2.8) +
  scale_fill_manual(values = stage_cols_lpg, breaks = ord_lpg, name = "") +
  theme_void(base_size = BASE) +
  labs(title = "b. Global VWA WTT CI share (%)") +
  geom_text(
    aes(label = percent(share, accuracy = 1)),
    position = position_stack(vjust = 0.7),
    size = pt_size(BASE * 0.9),
    colour = "black"
  ) +
  theme(
    legend.position = "none",
    plot.title = element_text(size = BASE, hjust = 0)
  )

donut_grob_lpg <- ggplotGrob(p_donut_lpg)

top_labels_lpg <- plot_df_lpg %>%
  slice_max(order_by = produced_volume_Mt, n = 15, with_ties = FALSE) %>%
  filter(!refinery %in% c("Jamnagar Gujarat (Reliance)", "Galveston Bay")) %>%
  pull(refinery)

top_labels_lpg <- unique(c(top_labels_lpg, "Ras Tanura"))

p_ref_lpg <- ggplot(plot_df_lpg) +
  geom_rect(aes(xmin = cumulative_prod_Mt - produced_volume_Mt, xmax = cumulative_prod_Mt,
                ymin = offset_0, ymax = offset_1, fill = "upstream"),
            colour = "black", linewidth = 0.01) +
  geom_rect(aes(xmin = cumulative_prod_Mt - produced_volume_Mt, xmax = cumulative_prod_Mt,
                ymin = offset_1, ymax = offset_2, fill = "transport"),
            colour = "black", linewidth = 0.01) +
  geom_rect(aes(xmin = cumulative_prod_Mt - produced_volume_Mt, xmax = cumulative_prod_Mt,
                ymin = offset_2, ymax = offset_3, fill = "refining"),
            colour = "black", linewidth = 0.01) +
  geom_rect(aes(xmin = cumulative_prod_Mt - produced_volume_Mt, xmax = cumulative_prod_Mt,
                ymin = offset_3, ymax = offset_4, fill = "distribution"),
            colour = "black", linewidth = 0.01) +
  scale_x_continuous(
    limits = c(0, max(plot_df_lpg$cumulative_prod_Mt, na.rm = TRUE)),
    breaks = c(0, 0.25, 0.5, 0.75, 0.999) * max(plot_df_lpg$cumulative_prod_Mt, na.rm = TRUE),
    labels = c(0, 25, 50, 75, 100),
    name = "Cumulative refinery-produced LPG before shipping (%, ~3.7 Mbpd, ~119 Mt y⁻¹)\nLPG density assumption (bbl/t ≈ 11.2 avg LPG blend)",
    expand = c(0.01, 0.01)
  ) +
  scale_y_continuous(
    limits = c(0, max(plot_df_lpg$offset_4, na.rm = TRUE) * 1.1),
    expand = c(0.02, 0.05),
    name = expression(atop("Refinery-level Well-to-Tank", "LPG Carbon Intensity (gCO"[2]~"e MJ"^-1*")"))
  ) +
  scale_fill_manual(values = stage_cols_lpg, breaks = ord_lpg, name = "") +
  labs(title = "A. Refinery-level LPG CI before export/import dynamics") +
  theme_marine(BASE) +
  theme(
    plot.title = element_text(size = BASE, hjust = 0),
    legend.position = c(0.02, 0.98),
    legend.justification = c(0, 1),
    legend.direction = "horizontal",
    legend.key.width = unit(1.2, "cm")
  ) +
  annotation_custom(
    bar_grob_lpg,
    xmin = max(plot_df_lpg$cumulative_prod_Mt, na.rm = TRUE) * 0.05,
    xmax = max(plot_df_lpg$cumulative_prod_Mt, na.rm = TRUE) * 0.45,
    ymin = max(plot_df_lpg$offset_4, na.rm = TRUE) * 0.30,
    ymax = max(plot_df_lpg$offset_4, na.rm = TRUE) * 0.96
  ) +
  annotation_custom(
    donut_grob_lpg,
    xmin = max(plot_df_lpg$cumulative_prod_Mt, na.rm = TRUE) * 0.55,
    xmax = max(plot_df_lpg$cumulative_prod_Mt, na.rm = TRUE) * 0.90,
    ymin = max(plot_df_lpg$offset_4, na.rm = TRUE) * 0.30,
    ymax = max(plot_df_lpg$offset_4, na.rm = TRUE) * 0.96
  ) +
  geom_hline(yintercept = global_avg_CI_lpg, linetype = "dashed", colour = "blue", linewidth = 0.8) +
  annotate(
    "text",
    x = max(plot_df_lpg$cumulative_prod_Mt, na.rm = TRUE) * 0.10,
    y = global_avg_CI_lpg + 7,
    label = paste0("Global VWA CI = ", sprintf("%.1f", global_avg_CI_lpg), " gCO₂e/MJ"),
    colour = "black",
    size = pt_size(BASE * 0.9),
    hjust = 0
  ) +
  geom_text(
    data = plot_df_lpg %>% filter(refinery %in% top_labels_lpg),
    aes(x = cumulative_prod_Mt - produced_volume_Mt / 2, y = 0, label = refinery),
    angle = 90,
    size = pt_size(BASE * 0.9),
    hjust = -0.02,
    vjust = 0.5,
    colour = "black"
  )

# ---------------------------------------------------------------------------
# Panel B
# ---------------------------------------------------------------------------
ci_asset <- lpg_clean_2 %>%
  transmute(
    country   = str_trim(country),
    vol_bbl_d = nz_num(lpg_volume_bbl_d),
    up_g_MJ   = nz_num(crude_intake_upstream_ci_g_co2e_mj),
    tr_g_MJ   = nz_num(crude_transport_ci_g_co2e_mj),
    ref_g_MJ  = nz_num(lpg_refining_ci_g_co2e_mj),
    dist_g_MJ = distribution_CI_gCO2e_MJ
  ) %>%
  mutate(total_ci_g_MJ = up_g_MJ + tr_g_MJ + ref_g_MJ + dist_g_MJ) %>%
  filter(!is.na(country), vol_bbl_d > 0, is.finite(total_ci_g_MJ))

top10_names <- ci_asset %>%
  group_by(country) %>%
  summarise(prod_bbl_d = sum(vol_bbl_d, na.rm = TRUE), .groups = "drop") %>%
  slice_max(prod_bbl_d, n = 10, with_ties = FALSE) %>%
  pull(country)

iso3_from_name <- function(x) {
  out <- countrycode(x, origin = "country.name", destination = "iso3c", warn = FALSE)
  ifelse(is.na(out) | out == "", x, out)
}

df2 <- ci_asset %>%
  filter(country %in% top10_names) %>%
  mutate(iso3 = iso3_from_name(country))

x_order <- df2 %>%
  group_by(iso3) %>%
  summarise(prod_bbl_d = sum(vol_bbl_d, na.rm = TRUE), .groups = "drop") %>%
  arrange(desc(prod_bbl_d)) %>%
  pull(iso3)

global_prod_totals <- ci_asset %>%
  group_by(country) %>%
  summarise(prod_bbl_d = sum(vol_bbl_d, na.rm = TRUE), .groups = "drop")

global_total_bbl_d <- sum(global_prod_totals$prod_bbl_d, na.rm = TRUE)

share_tbl <- global_prod_totals %>%
  filter(country %in% top10_names) %>%
  mutate(iso3 = iso3_from_name(country)) %>%
  group_by(iso3) %>%
  summarise(prod_bbl_d = sum(prod_bbl_d, na.rm = TRUE), .groups = "drop") %>%
  mutate(share = if (global_total_bbl_d > 0) prod_bbl_d / global_total_bbl_d else 0) %>%
  arrange(desc(share))

col_vec <- c(
  "#2BA6A6", "#E46C5A", "#5D6FEF", "#8F63A8", "#4AA84A",
  "#F39C3D", "#76B947", "#6F6BC1", "#2BA9CF", "#5A9BE6"
)
col_vec <- setNames(col_vec[seq_along(x_order)], x_order)

lab_vec <- setNames(
  paste0(x_order, " (", percent(share_tbl$share[match(x_order, share_tbl$iso3)], accuracy = 1), ")"),
  x_order
)

p_violin <- ggplot(
  df2,
  aes(x = factor(iso3, levels = x_order), y = total_ci_g_MJ, weight = vol_bbl_d, fill = iso3)
) +
  geom_violin(scale = "width", trim = TRUE, colour = "grey40", linewidth = 0.3) +
  stat_summary(fun = median, geom = "point", size = pt_size(BASE * 0.9), colour = "black") +
  scale_fill_manual(values = col_vec, breaks = x_order, labels = lab_vec, name = "Share of global LPG production") +
  labs(
    title = "B. Asset-level variability in LPG Well-to-Tank CI across the ten largest producing countries (73% cumulative production)",
    x = NULL,
    y = expression("Variable distribution of LPG WTT CI (gCO"[2]*"e MJ"^-1*")")
  ) +
  theme_bw(base_size = BASE) +
  theme(
    plot.title = element_text(hjust = 0, size = BASE),
    legend.position = c(0.02, 0.98),
    legend.justification = c(0, 1),
    legend.direction = "horizontal",
    legend.key.width = unit(1.2, "cm"),
    panel.grid.minor = element_blank()
  )

# ---------------------------------------------------------------------------
# Combine and save
# ---------------------------------------------------------------------------
combined <- (p_ref_lpg + theme(plot.margin = margin(8, 12, 6, 8))) /
  (p_violin  + theme(plot.margin = margin(6, 12, 8, 8))) +
  plot_layout(heights = c(1, 1), guides = "keep")

ggsave(file.path(fig_dir, "Fig 4.png"), combined, width = 12, height = 13, dpi = 300)
ggsave(file.path(fig_dir, "Fig 4.svg"), combined, width = 12, height = 13, dpi = 300)

# ---------------------------------------------------------------------------
# Excel bundle
# ---------------------------------------------------------------------------
inputs_manifest <- tibble(
  type = "input",
  path = c(file_path, inp_archie),
  detail = c("LPG PRELIM workbook", "Archie out_refineries")
)

outputs_manifest <- tibble(
  type = "df",
  name = c(
    "lpg_raw_clean", "lpg_clean", "lpg_clean_joined",
    "lpg_working_df", "ref_df_lpg", "plot_df_lpg",
    "global_dom_lpg", "violin_df_top10", "share_tbl"
  ),
  nrow = c(
    nrow(lpg), nrow(lpg_clean), nrow(lpg_clean_2),
    nrow(lpg_df), nrow(ref_df_lpg), nrow(plot_df_lpg),
    nrow(global_dom_lpg), nrow(df2), nrow(share_tbl)
  ),
  ncol = c(
    ncol(lpg), ncol(lpg_clean), ncol(lpg_clean_2),
    ncol(lpg_df), ncol(ref_df_lpg), ncol(plot_df_lpg),
    ncol(global_dom_lpg), ncol(df2), ncol(share_tbl)
  )
)

figs_manifest <- tibble(
  type = "figure",
  path = file.path(fig_dir, c("Fig 4.png", "Fig 4.svg"))
)

fig4_manifest <- bind_rows(inputs_manifest, outputs_manifest, figs_manifest)

sheet_list <- list(
  A_raw_asset_file_clean = lpg,
  A_lpg_clean = lpg_clean,
  A_lpg_clean_joined = lpg_clean_2,
  A_domestic_working_df = lpg_df,
  A_refinery_level_wavgs = ref_df_lpg,
  A_plot_df_stacked = plot_df_lpg,
  A_global_dom_wavgs = global_dom_lpg,
  B_top10_violin_df = df2,
  B_top10_global_shares = share_tbl
)

save_excel_bundle(
  sheet_list = sheet_list,
  manifest_df = fig4_manifest,
  out_xlsx = file.path(fig_dir, "Fig4_data_bundle.xlsx")
)
###############################################################################
# Fig. 5 - LNG
# Voyage-level stacked LNG WTT CI
###############################################################################

source("marine_figures_shared_helpers.R")

# ---------------------------------------------------------------------------
# User paths
# ---------------------------------------------------------------------------
fig_dir <- "PATH/Fig 5"
dir.create(fig_dir, recursive = TRUE, showWarnings = FALSE)

root_archie <- "PATH/Archie LNG v15"
lng_updated_dir <- "PATH/06 LNG model results"
inp_lngs  <- file.path(root_archie, "lngs.geojson")
inp_routes <- file.path(root_archie, "gasSeaRoutes-with-meta.geojson")
inp_regas <- file.path(root_archie, "regas-combined.geojson")
inp_lng_csv <- file.path(lng_updated_dir, "LNG_CI_calculations_MERGED_UPDATED.csv")

stopifnot(file.exists(inp_lngs), file.exists(inp_routes), file.exists(inp_regas), file.exists(inp_lng_csv))

# ---------------------------------------------------------------------------
# Read inputs
# ---------------------------------------------------------------------------
lngs <- sf::st_read(inp_lngs, quiet = TRUE)
gasSeaRoutes <- sf::st_read(inp_routes, quiet = TRUE)
regas <- sf::st_read(inp_regas, quiet = TRUE)
lng_1 <- read.csv(inp_lng_csv, stringsAsFactors = FALSE)

name_map <- c(
  "Bethioua" = "Arzew-Bethioua",
  "Skikda" = "Skikda",
  "APLNG" = "APLNG",
  "Darwin" = "Darwin",
  "GLNG" = "Gladstone",
  "Gorgon" = "Gorgon",
  "Ichthys" = "Ichthys",
  "NWS" = "North West Shelf",
  "Pluto" = "Pluto",
  "Prelude" = "Prelude FLNG",
  "QCLNG" = "QCLNG",
  "Wheatstone" = "Wheatstone",
  "Bontang" = "Bontang",
  "DSLNG" = "DSLNG",
  "Tangguh" = "Tangguh",
  "Bintulu" = "Bintulu",
  "PFLNG 1 Sabah" = "PFLNG 1 Sabah",
  "PFLNG 2" = "PFLNG 2",
  "Bonny LNG" = "Bonny",
  "Qalhat" = "Qalhat",
  "PNG LNG" = "PNG LNG",
  "Ras Laffan" = "Ras Laffan",
  "Sakhalin 2" = "Sakhalin II",
  "Yamal" = "Yamal",
  "Portovaya" = "Portovaya",
  "Vysotsk" = "Vysotsk",
  "Calcasieu Pass" = "Calcasieu Pass",
  "Cameron (Liqu.)" = "Cameron",
  "Corpus Christi" = "Corpus Christi",
  "Cove Point" = "Cove Point",
  "Elba Island Liq." = "Elba Island",
  "Freeport" = "Freeport Texas",
  "Sabine Pass" = "Sabine Pass",
  "Das Island" = "Das Island",
  "Atlantic LNG" = "Atlantic LNG",
  "Idku" = "Idku",
  "Peru LNG" = "Peru LNG",
  "Damietta" = "Damietta",
  "Cameroon FLNG Terminal" = "Cameroon FLNG Terminal",
  "Bioko" = "Bioko",
  "Soyo" = "Soyo",
  "Snohvit" = "Snohvit",
  "Lumut I" = "Lumut I"
)

lngs <- lngs %>%
  mutate(
    LNG.Name = str_squish(as.character(LNG.Name)),
    lng_archie = dplyr::recode(LNG.Name, !!!name_map, .default = LNG.Name)
  )

lngs_1 <- lngs %>%
  left_join(lng_1, by = "lng_archie")

num <- function(x) suppressWarnings(as.numeric(x))

lngs_2 <- lngs_1 %>%
  sf::st_drop_geometry() %>%
  mutate(
    country_clean = dplyr::coalesce(country, Country),
    iso3_clean = dplyr::coalesce(`Country.ISO.code`, countrycode(country_clean, "country.name", "iso3c"))
  ) %>%
  transmute(
    lng_id    = as.integer(`LNG.ID`),
    lng_name  = `LNG.Name`,
    country   = country_clean,
    iso3      = iso3_clean,
    export_tpy = num(`Exporting.LNG..tonne.year`),
    surfaceprocessing = num(`Surface.Processing..gCO2eq..MJ`),
    flaring           = num(`Flaring..gCO2eq..MJ`),
    venting           = num(`CO2.Venting...AGR..gCO2eq..MJ`),
    extraction        = num(`Production.and.Extraction..gCO2eq..MJ`),
    `methane super`   = num(`Methane.Leak...Super.Emitter...Ave..gCO2eq..MJ`),
    `methane other`   = num(`Methane.Leak...Other...Ave..gCO2eq..MJ`),
    drilling          = num(`Exploration.and.Drilling..gCO2eq..MJ`),
    offsite           = num(`Offsite..gCO2eq..MJ`),
    `ff to pp`        = num(`F2PP.Transmission..gCO2eq..MJ`),
    `pp to lp`        = num(`PP2L.Transmission..gCO2eq..MJ`),
    liquefaction      = num(Final_LNG_CI_gCO2e_per_MJ),
    lng_archie        = lng_archie
  ) %>%
  mutate(across(where(is.character), ~ na_if(., "")))

gasSeaRoutes_2 <- gasSeaRoutes %>%
  sf::st_drop_geometry() %>%
  select(-any_of(c("name", "route", "archie_origin_terminal_id", "archie_destination_terminal_id", "Description"))) %>%
  mutate(lng_name = origin)

gasSeaRoutes_2x <- gasSeaRoutes_2 %>%
  mutate(lng_key = norm_key(lng_name))

lngs_2x <- lngs_2 %>%
  mutate(lng_key = norm_key(lng_name))

gasSeaRoutes_3 <- gasSeaRoutes_2x %>%
  left_join(lngs_2x %>% select(-lng_name), by = "lng_key", suffix = c("", "_lng")) %>%
  mutate(lng_match_missing = if_else(is.na(lng_id), TRUE, FALSE))

gasSeaRoutes_31 <- gasSeaRoutes_3 %>%
  select(-lng_key, -export_tpy)

regas_1 <- regas %>%
  sf::st_drop_geometry() %>%
  select(
    Regas.ID, Regas.Name, Regas.Country, Regas.Country.ISO,
    Importing.LNG..tonne.year, Shipping..gCO2eq..MJ, Regasification..gCO2eq..MJ
  ) %>%
  rename(destination = Regas.Name)

regas_2 <- regas_1 %>%
  rename(
    regas_id = Regas.ID,
    regas_country = Regas.Country,
    regas_country_ISO = Regas.Country.ISO,
    import_tpy_chr = `Importing.LNG..tonne.year`,
    ship_ci_chr = `Shipping..gCO2eq..MJ`,
    regas_ci_chr = `Regasification..gCO2eq..MJ`
  ) %>%
  mutate(
    import_tpy = readr::parse_number(as.character(import_tpy_chr)),
    ship_ci_g_MJ = readr::parse_number(as.character(ship_ci_chr)),
    regas_ci_g_MJ = readr::parse_number(as.character(regas_ci_chr)),
    import_tpy_int = as.integer(round(import_tpy, 0))
  ) %>%
  select(regas_id, destination, regas_country, regas_country_ISO, import_tpy, import_tpy_int, ship_ci_g_MJ, regas_ci_g_MJ)

well_TO_regas <- gasSeaRoutes_31 %>%
  left_join(regas_2, by = "destination") %>%
  mutate(
    across(
      c(surfaceprocessing, flaring, venting, extraction, `methane super`, `methane other`,
        drilling, offsite, `ff to pp`, `pp to lp`, liquefaction, ship_ci_g_MJ, regas_ci_g_MJ),
      ~ suppressWarnings(as.numeric(.))
    )
  ) %>%
  mutate(
    WTL_ci_gco2e_mj = rowSums(across(c(
      surfaceprocessing, flaring, venting, extraction, `methane super`, `methane other`,
      drilling, offsite, `ff to pp`, `pp to lp`, liquefaction
    )), na.rm = TRUE),
    WTS_ci_gco2e_mj = rowSums(across(c(
      surfaceprocessing, flaring, venting, extraction, `methane super`, `methane other`,
      drilling, offsite, `ff to pp`, `pp to lp`, liquefaction, ship_ci_g_MJ
    )), na.rm = TRUE),
    WTR_ci_gco2e_mj = rowSums(across(c(
      surfaceprocessing, flaring, venting, extraction, `methane super`, `methane other`,
      drilling, offsite, `ff to pp`, `pp to lp`, liquefaction, ship_ci_g_MJ, regas_ci_g_MJ
    )), na.rm = TRUE),
    total_ci_gco2e_mj = WTR_ci_gco2e_mj
  )

# ---------------------------------------------------------------------------
# Volume for weighting
# ---------------------------------------------------------------------------
if (!"volume_bcm" %in% names(well_TO_regas)) {
  stop("The LNG route-level dataframe does not contain 'volume_bcm'. Add it upstream before running Fig. 5.")
}

wtr_raw <- well_TO_regas %>%
  mutate(
    volume_bcm = nz_num(volume_bcm),
    liquefaction_name = coalesce(.data[["lng_archie"]], .data[["lng_name"]]),
    regas_name = .data[["destination"]]
  ) %>%
  filter(surfaceprocessing > 0, flaring > 0, venting > 0, extraction > 0) %>%
  filter(is.finite(volume_bcm), volume_bcm > 0)

readr::write_csv(wtr_raw, file.path(fig_dir, "lng_complete_supply_chain.csv"))

stage_order <- c(
  "surfaceprocessing", "flaring", "venting", "extraction",
  "methane super", "methane other", "drilling", "offsite",
  "ff to pp", "pp to lp", "liquefaction", "ship_ci_g_MJ", "regas_ci_g_MJ"
)

stage_cols <- c(
  surfaceprocessing = "#80B1D3",
  flaring = "#FDB462",
  venting = "#B3DE69",
  extraction = "#BC80BD",
  `methane super` = "#FB8072",
  `methane other` = "#CCEBC5",
  drilling = "#8DD3C7",
  offsite = "#FFFFB3",
  `ff to pp` = "#BEBADA",
  `pp to lp` = "#FDBF6F",
  liquefaction = "#1F78B4",
  ship_ci_g_MJ = "#A6CEE3",
  regas_ci_g_MJ = "#33A02C"
)

display_lab <- c(
  surfaceprocessing = "surface processing",
  flaring = "flaring",
  venting = "venting (AGR)",
  extraction = "extraction",
  `methane super` = "methane super emitters",
  `methane other` = "methane (other)",
  drilling = "drilling",
  offsite = "offsite",
  `ff to pp` = "FF → PP",
  `pp to lp` = "PP → LP",
  liquefaction = "liquefaction",
  ship_ci_g_MJ = "shipping",
  regas_ci_g_MJ = "regasification"
)

voy <- wtr_raw %>%
  arrange(total_ci_gco2e_mj) %>%
  mutate(
    produced_volume_bcm = volume_bcm,
    cumulative_bcm = cumsum(produced_volume_bcm)
  )

offs <- voy %>% mutate(offset_0 = 0)
for (i in seq_along(stage_order)) {
  offs[[paste0("offset_", i)]] <- offs[[paste0("offset_", i - 1)]] + offs[[stage_order[i]]]
}

rect_df <- bind_rows(
  lapply(seq_along(stage_order), function(i) {
    tibble(
      cumulative_bcm = offs$cumulative_bcm,
      produced_volume_bcm = offs$produced_volume_bcm,
      xmin = offs$cumulative_bcm - offs$produced_volume_bcm,
      xmax = offs$cumulative_bcm,
      ymin = offs[[paste0("offset_", i - 1)]],
      ymax = offs[[paste0("offset_", i)]],
      stage = factor(stage_order[i], levels = stage_order)
    )
  })
)

max_x <- max(offs$cumulative_bcm, na.rm = TRUE)
max_y <- max(offs[[paste0("offset_", length(stage_order))]], na.rm = TRUE)

g_liq   <- wmean_na(voy$WTL_ci_gco2e_mj, voy$produced_volume_bcm)
g_ship  <- wmean_na(voy$WTS_ci_gco2e_mj, voy$produced_volume_bcm)
g_regas <- wmean_na(voy$WTR_ci_gco2e_mj, voy$produced_volume_bcm)

global_stage_vals <- tibble(
  stage = factor(stage_order, levels = stage_order),
  value = sapply(stage_order, function(s) wmean_na(voy[[s]], voy$produced_volume_bcm))
) %>%
  mutate(share = value / sum(value, na.rm = TRUE))

stage_order_bar <- global_stage_vals %>%
  arrange(desc(value)) %>%
  pull(stage) %>%
  as.character()

legend_labels <- setNames(
  paste0(display_lab[stage_order_bar], " (", percent(global_stage_vals$share[match(stage_order_bar, global_stage_vals$stage)], accuracy = 0.01), ")"),
  stage_order_bar
)

p_bar <- ggplot(global_stage_vals %>% mutate(stage = factor(stage, levels = stage_order_bar)),
                aes(stage, value, fill = stage)) +
  geom_col(width = 0.7, colour = "black", linewidth = 0.2) +
  scale_fill_manual(values = stage_cols[stage_order_bar], breaks = stage_order_bar, labels = display_lab[stage_order_bar], name = "") +
  scale_x_discrete(labels = display_lab[stage_order_bar]) +
  labs(
    x = "",
    y = expression(atop("Global VWA WTT CI", "(gCO"[2]*"e MJ"^-1*")")),
    title = "a. Well-to-Tank stages CI (high → low)"
  ) +
  theme_bw(base_size = BASE) +
  theme(
    axis.text.x = element_blank(),
    axis.ticks.x = element_blank(),
    legend.position = "none",
    plot.title = element_text(size = BASE)
  )

p_donut <- ggplot(global_stage_vals %>% mutate(stage = factor(stage, levels = stage_order_bar)),
                  aes(x = 2, y = share, fill = stage)) +
  geom_col(width = 1, colour = "black", linewidth = 0.2) +
  coord_polar(theta = "y") +
  xlim(1, 2.8) +
  scale_fill_manual(values = stage_cols[stage_order_bar], breaks = stage_order_bar, labels = legend_labels, name = "Stage share") +
  guides(fill = guide_legend(nrow = 3, byrow = TRUE)) +
  theme_void(base_size = BASE) +
  theme(legend.position = "top")

legend_g <- extract_legend_grob(p_donut)

p_donut2 <- ggplot(global_stage_vals %>% mutate(stage = factor(stage, levels = stage_order_bar)),
                   aes(x = 2, y = share, fill = stage)) +
  geom_col(width = 1, colour = "black", linewidth = 0.2) +
  coord_polar(theta = "y") +
  xlim(1, 2.8) +
  scale_fill_manual(values = stage_cols[stage_order_bar], breaks = stage_order_bar, labels = legend_labels, name = "Stage share") +
  theme_void(base_size = BASE) +
  labs(title = "b. Global WTT stage shares (%)") +
  theme(
    legend.position = "none",
    plot.title = element_text(size = BASE, hjust = 0.5),
    plot.margin = margin(5, 10, 5, 5)
  )

p_wtr <- ggplot(rect_df) +
  geom_rect(aes(xmin = xmin, xmax = xmax, ymin = ymin, ymax = ymax, fill = stage),
            colour = "black", linewidth = 0.015) +
  scale_fill_manual(values = stage_cols, breaks = stage_order, labels = display_lab[stage_order], name = "") +
  scale_x_continuous(
    limits = c(0, max_x * 1.005),
    breaks = c(0, 0.25 * max_x, 0.5 * max_x, 0.75 * max_x, max_x),
    labels = c(0, 25, 50, 75, 100),
    name = "Cumulative exported LNG (%, 17.1×10³ BCF, ~13.5 Mbpd, ~351 Mt y⁻¹)\nassumptions (1 Mtpa ~ 48.7 Bcf/yr; LNG density ~ 0.45 t/m³)",
    expand = c(0.01, 0.01)
  ) +
  scale_y_continuous(
    limits = c(0, max_y * 1.3),
    expand = c(0.02, 0.04),
    name = expression(atop("Voyage-level Well-to-Tank", "Carbon Intensity (gCO"[2]~"e MJ"^-1*")"))
  ) +
  geom_hline(yintercept = g_liq,   linetype = "dashed", colour = stage_cols[["liquefaction"]]) +
  geom_hline(yintercept = g_ship,  linetype = "dashed", colour = stage_cols[["ship_ci_g_MJ"]]) +
  geom_hline(yintercept = g_regas, linetype = "dashed", colour = stage_cols[["regas_ci_g_MJ"]]) +
  annotate("text", x = 0.02 * max_x, y = g_liq + 0.7,
           label = sprintf("Global VWA W-to-L CI = %.2f", g_liq),
           hjust = 0, size = pt_size(BASE * 0.9), colour = stage_cols[["liquefaction"]]) +
  annotate("text", x = 0.02 * max_x, y = g_ship + 0.7,
           label = sprintf("Global VWA W-to-S CI = %.2f", g_ship),
           hjust = 0, size = pt_size(BASE * 0.9), colour = stage_cols[["ship_ci_g_MJ"]]) +
  annotate("text", x = 0.30 * max_x, y = g_regas + 0.7,
           label = sprintf("Global VWA W-to-R CI = %.2f", g_regas),
           hjust = 0, size = pt_size(BASE * 0.9), colour = stage_cols[["regas_ci_g_MJ"]]) +
  labs(title = "Voyage-level LNG WTT CI (stacked by stage)") +
  theme_marine(BASE) +
  theme(
    plot.title = element_text(hjust = 0, size = BASE),
    legend.position = "none"
  ) +
  annotation_custom(
    ggplotGrob(p_bar),
    xmin = max_x * 0.05, xmax = max_x * 0.45,
    ymin = max_y * 0.65, ymax = max_y * 1.28
  ) +
  annotation_custom(
    ggplotGrob(p_donut2),
    xmin = max_x * 0.48, xmax = max_x * 0.85,
    ymin = max_y * 0.60, ymax = max_y * 1.28
  )

combo <- gridExtra::arrangeGrob(
  p_wtr,
  legend_g,
  ncol = 1,
  heights = unit.c(unit(1, "null"), grobHeight(legend_g) + unit(2, "mm"))
)

ggsave(file.path(fig_dir, "Fig 5.png"), combo, width = 12, height = 9, dpi = 300)
ggsave(file.path(fig_dir, "Fig 5.svg"), combo, width = 12, height = 9, dpi = 300)

# ---------------------------------------------------------------------------
# Excel bundle
# ---------------------------------------------------------------------------
stage_map <- tibble(
  stage_key = stage_order,
  stage_label = unname(display_lab[stage_order]),
  stage_col = unname(stage_cols[stage_order])
)

fig5_globals <- tibble(
  metric = c("g_liq_WTL", "g_ship_WTS", "g_regas_WTR", "max_x_cumulative_bcm", "max_y_stack"),
  value  = c(g_liq, g_ship, g_regas, max_x, max_y)
)

fig5_legend <- tibble(
  stage_key = stage_order_bar,
  stage_label = unname(display_lab[stage_order_bar]),
  legend_label = unname(legend_labels[stage_order_bar]),
  share = global_stage_vals$share[match(stage_order_bar, global_stage_vals$stage)],
  value_gco2e_mj = global_stage_vals$value[match(stage_order_bar, global_stage_vals$stage)]
)

wtr_raw_fig5 <- wtr_raw %>%
  select(
    any_of(c("liquefaction_name", "regas_name", "destination", "lng_archie", "lng_name")),
    any_of(c("volume_bcm", "produced_volume_bcm", "cumulative_bcm", "total_ci_gco2e_mj", "WTL_ci_gco2e_mj", "WTS_ci_gco2e_mj", "WTR_ci_gco2e_mj")),
    all_of(intersect(stage_order, names(.)))
  )

voy_fig5 <- voy %>%
  select(
    any_of(c("liquefaction_name", "regas_name", "destination", "lng_archie", "lng_name")),
    volume_bcm, produced_volume_bcm, cumulative_bcm, total_ci_gco2e_mj,
    WTL_ci_gco2e_mj, WTS_ci_gco2e_mj, WTR_ci_gco2e_mj,
    all_of(intersect(stage_order, names(.)))
  )

rect_df_fig5 <- rect_df %>%
  mutate(
    stage_key = as.character(stage),
    stage_label = unname(display_lab[stage_key]),
    stage_col = unname(stage_cols[stage_key])
  ) %>%
  select(cumulative_bcm, produced_volume_bcm, xmin, xmax, ymin, ymax, stage_key, stage_label, stage_col)

global_stage_vals_fig5 <- global_stage_vals %>%
  mutate(
    stage_key = as.character(stage),
    stage_label = unname(display_lab[stage_key]),
    stage_col = unname(stage_cols[stage_key])
  ) %>%
  select(stage_key, stage_label, stage_col, value, share)

outputs_manifest <- tibble(
  type = "df",
  name = c("stage_map", "fig5_globals", "fig5_legend", "wtr_raw_fig5", "voy_fig5", "rect_df_fig5", "global_stage_vals_fig5"),
  nrow = c(nrow(stage_map), nrow(fig5_globals), nrow(fig5_legend), nrow(wtr_raw_fig5), nrow(voy_fig5), nrow(rect_df_fig5), nrow(global_stage_vals_fig5)),
  ncol = c(ncol(stage_map), ncol(fig5_globals), ncol(fig5_legend), ncol(wtr_raw_fig5), ncol(voy_fig5), ncol(rect_df_fig5), ncol(global_stage_vals_fig5))
)

figs_manifest <- tibble(
  type = "figure",
  path = file.path(fig_dir, c("Fig 5.png", "Fig 5.svg"))
)

fig5_manifest <- bind_rows(outputs_manifest, figs_manifest)

sheet_list <- list(
  stage_map = stage_map,
  globals_used_in_plot = fig5_globals,
  legend_and_bar_order = fig5_legend,
  A_wtr_raw_min_for_fig5 = wtr_raw_fig5,
  B_voy_filtered_ordered = voy_fig5,
  C_rect_df_stacked_geom = rect_df_fig5,
  D_global_stage_vals = global_stage_vals_fig5
)

save_excel_bundle(
  sheet_list = sheet_list,
  manifest_df = fig5_manifest,
  out_xlsx = file.path(fig_dir, "Fig5_data_bundle.xlsx")
)

###############################################################################
# Fig. 7 - comparison of HSFO, LPG, LNG
# Requires final WTW-ready dataframes for each fuel
###############################################################################

source("marine_figures_shared_helpers.R")

# ---------------------------------------------------------------------------
# User paths
# Replace these with your final files
# ---------------------------------------------------------------------------
fig_dir <- "PATH/Fig 7"
dir.create(fig_dir, recursive = TRUE, showWarnings = FALSE)

inp_hsfo <- "D:/PATH/TO/hsfo_df_c.csv"
inp_lpg  <- "D:/PATH/TO/lpg_df3.csv"
inp_lng  <- "D:/PATH/TO/lng_df2.csv"

stopifnot(file.exists(inp_hsfo), file.exists(inp_lpg), file.exists(inp_lng))

hsfo_df_c <- readr::read_csv(inp_hsfo, show_col_types = FALSE)
lpg_df3   <- readr::read_csv(inp_lpg, show_col_types = FALSE)
lng_df2   <- readr::read_csv(inp_lng, show_col_types = FALSE)

BASE <- 12

# ---------------------------------------------------------------------------
# Panel A
# ---------------------------------------------------------------------------
hsfo_p1 <- hsfo_df_c %>%
  filter(WTW_CI_gCO2e_MJ < 114) %>%
  select(WTW_CI_gCO2e_MJ, produced_volume_Mt, cumulative_prod_Mt) %>%
  arrange(WTW_CI_gCO2e_MJ, cumulative_prod_Mt) %>%
  mutate(share = cumulative_prod_Mt / sum(produced_volume_Mt))

lpg_p1 <- lpg_df3 %>%
  filter(WTW_CI_gCO2e_MJ < 150) %>%
  select(WTW_CI_gCO2e_MJ, produced_volume_Mt, cumulative_prod_Mt) %>%
  arrange(WTW_CI_gCO2e_MJ, cumulative_prod_Mt) %>%
  mutate(share = cumulative_prod_Mt / sum(produced_volume_Mt))

lng_p1 <- lng_df2 %>%
  select(WTW_CI_gCO2e_MJ, produced_volume_bcm, cumulative_bcm) %>%
  arrange(WTW_CI_gCO2e_MJ, cumulative_bcm) %>%
  mutate(share = cumulative_bcm / sum(produced_volume_bcm))

df_lines <- bind_rows(
  hsfo_p1 %>% transmute(fuel = "HSFO", share = share, WTW_CI_gCO2e_MJ = WTW_CI_gCO2e_MJ),
  lpg_p1  %>% transmute(fuel = "LPG",  share = share, WTW_CI_gCO2e_MJ = WTW_CI_gCO2e_MJ),
  lng_p1  %>% transmute(fuel = "LNG",  share = share, WTW_CI_gCO2e_MJ = WTW_CI_gCO2e_MJ)
) %>%
  group_by(fuel) %>%
  arrange(share, .by_group = TRUE) %>%
  ungroup()

vwa_hsfo <- hsfo_p1 %>%
  summarise(VWA_WTW = sum(WTW_CI_gCO2e_MJ * produced_volume_Mt, na.rm = TRUE) / sum(produced_volume_Mt, na.rm = TRUE)) %>%
  pull(VWA_WTW)

vwa_lpg <- lpg_p1 %>%
  summarise(VWA_WTW = sum(WTW_CI_gCO2e_MJ * produced_volume_Mt, na.rm = TRUE) / sum(produced_volume_Mt, na.rm = TRUE)) %>%
  pull(VWA_WTW)

vwa_lng <- lng_p1 %>%
  summarise(VWA_WTW = sum(WTW_CI_gCO2e_MJ * produced_volume_bcm, na.rm = TRUE) / sum(produced_volume_bcm, na.rm = TRUE)) %>%
  pull(VWA_WTW)

avg_df <- tibble(
  fuel = c("HSFO", "LPG", "LNG"),
  y = c(vwa_hsfo, vwa_lpg, vwa_lng)
)

p_fuels <- ggplot(df_lines, aes(x = share, y = WTW_CI_gCO2e_MJ, colour = fuel)) +
  geom_line(linewidth = 1.5) +
  geom_hline(data = avg_df, aes(yintercept = y, colour = fuel), linetype = "dashed", linewidth = 0.9, show.legend = FALSE) +
  geom_text(
    data = avg_df,
    aes(x = 0.30, y = y + 0.5, label = sprintf("%s VWA = %.2f", fuel, y), colour = fuel),
    hjust = 1,
    vjust = -0.2,
    size = pt_size(BASE * 0.9),
    show.legend = FALSE
  ) +
  scale_x_continuous(labels = percent_format(accuracy = 1), limits = c(0, 1), expand = c(0.01, 0.01)) +
  scale_color_npg() +
  labs(
    x = "Cumulative share of volume",
    y = expression("Well-To-Wake carbon intensity (gCO"[2]*"e MJ"^-1*")"),
    colour = NULL,
    title = "a. Comparison"
  ) +
  theme_marine(BASE) +
  theme(
    legend.position = c(0.02, 0.98),
    legend.justification = c(0, 1),
    legend.direction = "horizontal"
  )

# ---------------------------------------------------------------------------
# Panels B-G
# ---------------------------------------------------------------------------
top5_hsfo_ci <- hsfo_df_c %>%
  select(trade_countries, VWA_WTW_CI_gCO2e_MJ_trade) %>%
  filter(!is.na(trade_countries)) %>%
  group_by(trade_countries) %>%
  summarise(VWA_WTW_CI_gCO2e_MJ_trade = mean(VWA_WTW_CI_gCO2e_MJ_trade, na.rm = TRUE), .groups = "drop") %>%
  filter(is.finite(VWA_WTW_CI_gCO2e_MJ_trade)) %>%
  filter(trade_countries %in% c("ESP→NGA", "GRC→CHN", "DEU→FRA", "ESP→ARG", "NLD→AUS", "ESP→PAN", "MEX→SGP"))

bottom4_hsfo_ci <- hsfo_df_c %>%
  select(trade_countries, VWA_WTW_CI_gCO2e_MJ_trade) %>%
  filter(!is.na(trade_countries)) %>%
  group_by(trade_countries) %>%
  summarise(VWA_WTW_CI_gCO2e_MJ_trade = mean(VWA_WTW_CI_gCO2e_MJ_trade, na.rm = TRUE), .groups = "drop") %>%
  filter(is.finite(VWA_WTW_CI_gCO2e_MJ_trade)) %>%
  filter(trade_countries %in% c("SAU→CHN", "BHR→SGP", "FIN→DEU", "SWE→NOR", "DNK→EGY", "KWT→QAT", "BRA→URY"))

top5_lpg_ci <- lpg_df3 %>%
  select(country_code, VWA_WTW_CI_gCO2e_MJ_ref) %>%
  filter(!is.na(country_code)) %>%
  group_by(country_code) %>%
  summarise(VWA_WTW_CI_gCO2e_MJ_ref = mean(VWA_WTW_CI_gCO2e_MJ_ref, na.rm = TRUE), .groups = "drop") %>%
  filter(is.finite(VWA_WTW_CI_gCO2e_MJ_ref)) %>%
  filter(country_code %in% c("MEX", "SVK", "SRB", "AGO", "IND", "IDN", "BGR"))

bottom4_lpg_ci <- lpg_df3 %>%
  select(country_code, VWA_WTW_CI_gCO2e_MJ_ref) %>%
  filter(!is.na(country_code)) %>%
  group_by(country_code) %>%
  summarise(VWA_WTW_CI_gCO2e_MJ_ref = mean(VWA_WTW_CI_gCO2e_MJ_ref, na.rm = TRUE), .groups = "drop") %>%
  filter(is.finite(VWA_WTW_CI_gCO2e_MJ_ref)) %>%
  filter(country_code %in% c("BHR", "NOR", "SWE", "JOR", "PAK", "DNK", "SAU"))

top5_lng_ci <- lng_df2 %>%
  select(trade_countries, VWA_WTW_CI_gCO2e_MJ_trade) %>%
  filter(!is.na(trade_countries)) %>%
  group_by(trade_countries) %>%
  summarise(VWA_WTW_CI_gCO2e_MJ_trade = mean(VWA_WTW_CI_gCO2e_MJ_trade, na.rm = TRUE), .groups = "drop") %>%
  filter(is.finite(VWA_WTW_CI_gCO2e_MJ_trade)) %>%
  filter(trade_countries %in% c("IDN→MEX", "NGA→PRI", "MYS→KOR", "NGA→POL", "IDN→CHN", "AUS→MYS", "DZA→ARG"))

bottom4_lng_ci <- lng_df2 %>%
  select(trade_countries, VWA_WTW_CI_gCO2e_MJ_trade) %>%
  filter(!is.na(trade_countries)) %>%
  group_by(trade_countries) %>%
  summarise(VWA_WTW_CI_gCO2e_MJ_trade = mean(VWA_WTW_CI_gCO2e_MJ_trade, na.rm = TRUE), .groups = "drop") %>%
  filter(is.finite(VWA_WTW_CI_gCO2e_MJ_trade)) %>%
  filter(trade_countries %in% c("QAT→ARE", "AUS→IDN", "USA→MEX", "IDN→FRA", "RUS→IDN", "PNG→JPN", "DZA→FRA"))

make_bar_panel <- function(df, xcol, ycol, fill_col, title_txt, ylab_expr) {
  ggplot(df %>% mutate(.x = reorder(.data[[xcol]], .data[[ycol]])), aes(x = .x, y = .data[[ycol]])) +
    geom_col(fill = fill_col, colour = "black", linewidth = 0.2) +
    geom_text(aes(label = sprintf("%.1f", .data[[ycol]])), vjust = -0.25, hjust = 0.7, size = BASE / 3) +
    scale_y_continuous(
      limits = c(0, 120),
      breaks = c(0, 30, 60, 90, 120),
      labels = label_number(accuracy = 1),
      expand = expansion(mult = c(0.02, 0.08))
    ) +
    scale_x_discrete(guide = guide_axis(n.dodge = 2)) +
    labs(x = NULL, y = ylab_expr, title = title_txt) +
    theme_marine(BASE) +
    theme(legend.position = "none")
}

p_top5 <- make_bar_panel(top5_hsfo_ci, "trade_countries", "VWA_WTW_CI_gCO2e_MJ_trade", "#E46C5A", "b. Top HSFO",
                         expression("Trade VWA Well-To-Wake CI (gCO"[2]*"e MJ"^-1*")"))
p_bottom4 <- make_bar_panel(bottom4_hsfo_ci, "trade_countries", "VWA_WTW_CI_gCO2e_MJ_trade", "#2C7FB8", "e. Bottom HSFO",
                            expression("Trade VWA Well-To-Wake CI (gCO"[2]*"e MJ"^-1*")"))

p_top5_lpg <- make_bar_panel(top5_lpg_ci, "country_code", "VWA_WTW_CI_gCO2e_MJ_ref", "#6F6BC1", "c. Top ref-sourced LPG",
                             expression("Country VWA Well-To-Wake CI (gCO"[2]*"e MJ"^-1*")"))
p_bottom4_lpg <- make_bar_panel(bottom4_lpg_ci, "country_code", "VWA_WTW_CI_gCO2e_MJ_ref", "#F39C3D", "f. Bottom ref-sourced LPG",
                                expression("Trade VWA Well-To-Wake CI (gCO"[2]*"e MJ"^-1*")"))

p_top5_lng <- make_bar_panel(top5_lng_ci, "trade_countries", "VWA_WTW_CI_gCO2e_MJ_trade", "#4AA84A", "d. Top LNG",
                             expression("Country VWA Well-To-Wake CI (gCO"[2]*"e MJ"^-1*")"))
p_bottom4_lng <- make_bar_panel(bottom4_lng_ci, "trade_countries", "VWA_WTW_CI_gCO2e_MJ_trade", "#BC4E9C", "g. Bottom LNG",
                                expression("Trade VWA Well-To-Wake CI (gCO"[2]*"e MJ"^-1*")"))

row2 <- p_top5 | p_top5_lpg | p_top5_lng
row3 <- p_bottom4 | p_bottom4_lpg | p_bottom4_lng

combined <- p_fuels / row2 / row3 +
  plot_layout(heights = c(1, 1, 1), guides = "keep") &
  theme(plot.margin = margin(6, 10, 6, 10))

ggsave(file.path(fig_dir, "fuels_panel_a_bcd_efg.png"), combined, width = 13, height = 12, dpi = 300)
ggsave(file.path(fig_dir, "fuels_panel_a_bcd_efg.svg"), combined, width = 13, height = 12)

# ---------------------------------------------------------------------------
# Excel bundle
# ---------------------------------------------------------------------------
manifest <- tibble(
  type = c("input", "input", "input", "figure", "figure"),
  path = c(inp_hsfo, inp_lpg, inp_lng,
           file.path(fig_dir, "fuels_panel_a_bcd_efg.png"),
           file.path(fig_dir, "fuels_panel_a_bcd_efg.svg"))
)

sheet_list <- list(
  hsfo_p1 = hsfo_p1,
  lpg_p1 = lpg_p1,
  lng_p1 = lng_p1,
  df_lines = df_lines,
  avg_df = avg_df,
  top5_hsfo_ci = top5_hsfo_ci,
  bottom4_hsfo_ci = bottom4_hsfo_ci,
  top5_lpg_ci = top5_lpg_ci,
  bottom4_lpg_ci = bottom4_lpg_ci,
  top5_lng_ci = top5_lng_ci,
  bottom4_lng_ci = bottom4_lng_ci
)

save_excel_bundle(
  sheet_list = sheet_list,
  manifest_df = manifest,
  out_xlsx = file.path(fig_dir, "Fig7_data_bundle.xlsx")
)
