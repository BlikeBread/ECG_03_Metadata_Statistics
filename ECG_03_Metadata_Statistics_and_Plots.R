###############################################################################
#
# PURPOSE (Script 03/03):
#   Merge an upstream summary table with external metadata, then run group-wise
#   statistics + generate per-metric plots.
#
# INPUTS (GUI-selected):
#   1) Summary_FirstXXX_Final_2min.xlsx   (produced upstream; must contain mouse_id)
#   2) Metadata.xlsx                      (created externally; must contain mouse_id + Group)
#
# USER INTERFACE:
#   - Select which Group levels to include (plots + stats)
#   - Choose group order for plotting
#
# OUTPUTS (inside chosen OUTPUT folder):
#   03_Metadata_Analysis/
#     ├── merged_data_ALL.csv
#     ├── merged_data_FILTERED.csv
#     ├── stats_per_metric.csv                  (t-test if 2 groups; ANOVA if >2; empty if <2)
#     ├── stats_tukeyHSD_all_metrics.csv         (only if >2 groups)
#     └── Plots/*.pdf
#
# NOTES:
#   - Uses a local mean±SD summary function (no ggplot2::mean_sd dependency).
#   - Applies Tk option DB listbox theme to avoid "black listbox" on macOS.
###############################################################################

suppressPackageStartupMessages({ library(tcltk) })

ensure_pkg <- function(pkg) {
  if (!requireNamespace(pkg, quietly = TRUE)) install.packages(pkg, dependencies = TRUE)
}
pkgs <- c("dplyr","tidyr","ggplot2","openxlsx","tibble","stringr")
invisible(lapply(pkgs, ensure_pkg))

suppressPackageStartupMessages({
  library(tcltk)
  library(dplyr)
  library(tidyr)
  library(ggplot2)
  library(openxlsx)
  library(tibble)
  library(stringr)
})

# ---------------------------- Tk theme fix (IMPORTANT) ------------------------
# Fixes "black listbox / invisible text" especially on macOS Tk/Aqua
set_listbox_theme <- function() {
  try(tcl("option", "add", "*Listbox.background",       "white"), silent = TRUE)
  try(tcl("option", "add", "*Listbox.foreground",       "black"), silent = TRUE)
  try(tcl("option", "add", "*Listbox.selectBackground", "#cce5ff"), silent = TRUE)
  try(tcl("option", "add", "*Listbox.selectForeground", "black"), silent = TRUE)
  try(tcl("option", "add", "*listbox.background",       "white"), silent = TRUE)
  try(tcl("option", "add", "*listbox.foreground",       "black"), silent = TRUE)
  try(tcl("option", "add", "*listbox.selectBackground", "#cce5ff"), silent = TRUE)
  try(tcl("option", "add", "*listbox.selectForeground", "black"), silent = TRUE)
}
set_listbox_theme()

# ---------------------------- Plot summary helper -----------------------------
# Local mean ± SD function (avoids ggplot2::mean_sd which may not be exported)
mean_sd_local <- function(x, na.rm = TRUE) {
  x <- if (na.rm) x[is.finite(x)] else x
  m <- mean(x, na.rm = na.rm)
  s <- stats::sd(x, na.rm = na.rm)
  c(y = m, ymin = m - s, ymax = m + s)
}

# ---------------------------- GUI helpers ------------------------------------

pick_file <- function(caption) {
  f <- tcltk::tk_choose.files(caption = caption)
  if (!length(f) || is.na(f[1]) || !nzchar(f[1])) stop("No file selected.")
  normalizePath(f[1], winslash = "/", mustWork = TRUE)
}

pick_dir <- function(caption) {
  d <- tcltk::tk_choose.dir(caption = caption)
  if (is.na(d) || !nzchar(d)) stop("No folder selected.")
  normalizePath(d, winslash = "/", mustWork = TRUE)
}

gui_select_one <- function(title, choices, width_chars = 60, height_rows = 14) {
  if (!length(choices)) stop("No choices provided.")
  
  tt <- tktoplevel()
  tkwm.title(tt, title)
  tkwm.resizable(tt, 0, 0)
  
  res  <- tclVar("")
  done <- tclVar(0)
  
  frm <- tkframe(tt)
  lb <- tklistbox(frm, selectmode = "single", height = height_rows, width = width_chars, exportselection = FALSE)
  
  tkconfigure(lb,
              background       = "white",
              foreground       = "black",
              selectbackground = "#cce5ff",
              selectforeground = "black")
  
  sb <- tkscrollbar(frm, orient = "vertical", command = function(...) tkset(lb, ...))
  tkconfigure(lb, yscrollcommand = function(...) tkset(sb, ...))
  
  for (ch in choices) tkinsert(lb, "end", ch)
  
  tkgrid(lb, sb, sticky = "nsew")
  tkgrid.configure(sb, sticky = "ns")
  
  btn_frm <- tkframe(tt)
  
  on_ok <- function() {
    sel <- as.integer(tkcurselection(lb))
    tclvalue(res) <- if (length(sel) && !is.na(sel)) choices[sel + 1] else ""
    tclvalue(done) <- 1
  }
  on_cancel <- function() {
    tclvalue(res) <- ""
    tclvalue(done) <- 1
  }
  
  ok_btn <- tkbutton(btn_frm, text = "OK", width = 10, command = on_ok)
  cancel_btn <- tkbutton(btn_frm, text = "Cancel", width = 10, command = on_cancel)
  
  tkbind(lb, "<Double-Button-1>", function() on_ok())
  tkbind(tt, "<Escape>", function() on_cancel())
  tkbind(tt, "<Destroy>", function() {
    if (as.integer(tclvalue(done)) == 0) {
      tclvalue(res) <- ""
      tclvalue(done) <- 1
    }
  })
  
  tkgrid(frm, padx = 10, pady = c(10, 5))
  tkgrid(btn_frm, padx = 10, pady = c(0, 10), sticky = "e")
  tkgrid(ok_btn, cancel_btn, padx = 5)
  
  tkfocus(lb)
  suppressWarnings(try(tkgrab.set(tt), silent = TRUE))
  tkwait.variable(done)
  suppressWarnings(try(tkgrab.release(tt), silent = TRUE))
  suppressWarnings(try(tkdestroy(tt), silent = TRUE))
  
  out <- tclvalue(res)
  if (!nzchar(out)) stop("No selection made.")
  out
}

gui_select_many <- function(title, choices, width_chars = 60, height_rows = 16) {
  if (!length(choices)) stop("No choices provided.")
  
  tt <- tktoplevel()
  tkwm.title(tt, title)
  tkwm.resizable(tt, 0, 0)
  
  done <- tclVar(0)
  res  <- tclVar("")
  
  frm <- tkframe(tt)
  lb <- tklistbox(frm, selectmode = "extended",
                  height = min(height_rows, max(6, length(choices))),
                  width = width_chars, exportselection = FALSE)
  
  tkconfigure(lb,
              background       = "white",
              foreground       = "black",
              selectbackground = "#cce5ff",
              selectforeground = "black")
  
  sb <- tkscrollbar(frm, orient = "vertical", command = function(...) tkset(lb, ...))
  tkconfigure(lb, yscrollcommand = function(...) tkset(sb, ...))
  
  for (ch in choices) tkinsert(lb, "end", ch)
  
  tkgrid(lb, sb, sticky = "nsew")
  tkgrid.configure(sb, sticky = "ns")
  
  hint <- tklabel(tt, text = "Tip: Ctrl/Shift-click to select multiple groups.")
  btn_frm <- tkframe(tt)
  
  on_all <- function() tkselection.set(lb, 0, "end")
  on_none <- function() tkselection.clear(lb, 0, "end")
  
  on_ok <- function() {
    sel <- as.integer(tkcurselection(lb))
    if (length(sel) && !anyNA(sel)) {
      picked <- choices[sel + 1]
      tclvalue(res) <- paste(picked, collapse = "|||")
    } else {
      tclvalue(res) <- ""
    }
    tclvalue(done) <- 1
  }
  
  on_cancel <- function() {
    tclvalue(res) <- ""
    tclvalue(done) <- 1
  }
  
  all_btn    <- tkbutton(btn_frm, text = "Select all", width = 12, command = on_all)
  none_btn   <- tkbutton(btn_frm, text = "Clear",      width = 10, command = on_none)
  ok_btn     <- tkbutton(btn_frm, text = "OK",         width = 10, command = on_ok)
  cancel_btn <- tkbutton(btn_frm, text = "Cancel",     width = 10, command = on_cancel)
  
  tkbind(tt, "<Escape>", function() on_cancel())
  tkbind(tt, "<Destroy>", function() {
    if (as.integer(tclvalue(done)) == 0) {
      tclvalue(res) <- ""
      tclvalue(done) <- 1
    }
  })
  
  tkgrid(frm, padx = 10, pady = c(10, 5))
  tkgrid(hint, padx = 10, pady = c(0, 5), sticky = "w")
  tkgrid(btn_frm, padx = 10, pady = c(0, 10), sticky = "e")
  tkgrid(all_btn, none_btn, ok_btn, cancel_btn, padx = 5)
  
  tkfocus(lb)
  suppressWarnings(try(tkgrab.set(tt), silent = TRUE))
  tkwait.variable(done)
  suppressWarnings(try(tkgrab.release(tt), silent = TRUE))
  suppressWarnings(try(tkdestroy(tt), silent = TRUE))
  
  out <- tclvalue(res)
  if (!nzchar(out)) stop("No groups selected.")
  strsplit(out, "\\|\\|\\|", fixed = FALSE)[[1]]
}

# ---------------------------- Data utilities ---------------------------------

to_num <- function(x) suppressWarnings(as.numeric(gsub(",", ".", as.character(x))))

coerce_numeric_like <- function(df, keep_cols = c("mouse_id", "Group")) {
  for (nm in names(df)) {
    if (nm %in% keep_cols) next
    if (is.numeric(df[[nm]])) next
    vn <- to_num(df[[nm]])
    frac_ok <- mean(is.finite(vn), na.rm = TRUE)
    if (isTRUE(frac_ok >= 0.7)) df[[nm]] <- vn
  }
  df
}

# ---------------------------- 1) Select inputs --------------------------------

message("Select SUMMARY file produced upstream (must contain mouse_id).")
summary_file <- pick_file("Select Summary_FirstXX_Final_2min.xlsx")

message("Select METADATA file (must contain mouse_id + Group).")
metadata_file <- pick_file("Select Metadata.xlsx")

message("Select OUTPUT folder for Script 03 results.")
out_root <- pick_dir("Select output folder")

out_dir  <- file.path(out_root, "03_Metadata_Analysis")
plot_dir <- file.path(out_dir, "Plots")
dir.create(plot_dir, showWarnings = FALSE, recursive = TRUE)

# ---------------------------- 2) Load data ------------------------------------

summary_df  <- openxlsx::read.xlsx(summary_file)
metadata_df <- openxlsx::read.xlsx(metadata_file)

if (!("mouse_id" %in% names(summary_df))) stop("Summary file must contain a 'mouse_id' column.")
if (!all(c("mouse_id","Group") %in% names(metadata_df))) stop("Metadata must contain columns: mouse_id, Group")

summary_df$mouse_id  <- trimws(as.character(summary_df$mouse_id))
metadata_df$mouse_id <- trimws(as.character(metadata_df$mouse_id))
metadata_df$Group    <- trimws(as.character(metadata_df$Group))

summary_df <- coerce_numeric_like(summary_df, keep_cols = c("mouse_id"))

# ---------------------------- 3) Merge ----------------------------------------

merged_all <- summary_df %>%
  inner_join(metadata_df %>% select(mouse_id, Group), by = "mouse_id")

if (!nrow(merged_all)) stop("Merge resulted in 0 rows. Check that mouse_id values match between summary and metadata.")

write.csv(merged_all, file.path(out_dir, "merged_data_ALL.csv"), row.names = FALSE)

# ---------------------------- 3b) Choose groups -------------------------------

all_groups <- sort(unique(merged_all$Group))
selected_groups <- gui_select_many(
  title = "Select Group(s) to include (plots + stats)",
  choices = all_groups,
  width_chars = 50,
  height_rows = 14
)

order_mode <- gui_select_one(
  "Group order for plots",
  c("Use metadata order (alphabetical)", "Use selected order (as picked)"),
  width_chars = 45,
  height_rows = 6
)

merged_df <- merged_all %>% filter(Group %in% selected_groups)
if (!nrow(merged_df)) stop("After filtering by selected groups, 0 rows remain.")

if (order_mode == "Use selected order (as picked)") {
  merged_df$Group <- factor(merged_df$Group, levels = selected_groups)
} else {
  merged_df$Group <- factor(merged_df$Group, levels = sort(unique(merged_df$Group)))
}

write.csv(merged_df, file.path(out_dir, "merged_data_FILTERED.csv"), row.names = FALSE)
message("Included groups: ", paste(levels(merged_df$Group), collapse = ", "))

# ---------------------------- 4) Long format ---------------------------------

long_df <- merged_df %>%
  select(Group, where(is.numeric)) %>%
  pivot_longer(cols = -Group, names_to = "Metric", values_to = "Value") %>%
  filter(is.finite(Value))

n_groups <- nlevels(long_df$Group)
if (n_groups < 1) stop("No groups left after filtering.")
message(sprintf("Groups included: %d (%s)", n_groups, paste(levels(long_df$Group), collapse = ", ")))

# ---------------------------- 5) Statistics ----------------------------------

stats_df <- tibble()
tukey_df <- tibble()

if (n_groups >= 2) {
  stats_list <- list()
  tukey_list <- list()
  
  for (m in unique(long_df$Metric)) {
    df_m <- long_df %>% filter(Metric == m)
    if (nlevels(droplevels(df_m$Group)) < 2) next
    
    means_tbl <- df_m %>%
      group_by(Group) %>%
      summarise(
        Mean = mean(Value, na.rm = TRUE),
        SD   = sd(Value, na.rm = TRUE),
        N    = sum(is.finite(Value)),
        .groups = "drop"
      )
    
    if (n_groups == 2) {
      tt <- tryCatch(t.test(Value ~ Group, data = df_m), error = function(e) NULL)
      if (is.null(tt)) next
      glev <- levels(df_m$Group)
      
      stats_list[[m]] <- tibble(
        Metric = m,
        Test = "t-test",
        Group1 = glev[1],
        Group2 = glev[2],
        Mean_Group1 = means_tbl$Mean[means_tbl$Group == glev[1]],
        Mean_Group2 = means_tbl$Mean[means_tbl$Group == glev[2]],
        SD_Group1   = means_tbl$SD[means_tbl$Group == glev[1]],
        SD_Group2   = means_tbl$SD[means_tbl$Group == glev[2]],
        N_Group1    = means_tbl$N[means_tbl$Group == glev[1]],
        N_Group2    = means_tbl$N[means_tbl$Group == glev[2]],
        P_value = tt$p.value
      )
    } else {
      fit <- tryCatch(aov(Value ~ Group, data = df_m), error = function(e) NULL)
      if (is.null(fit)) next
      
      p_anova <- tryCatch(summary(fit)[[1]][["Pr(>F)"]][1], error = function(e) NA_real_)
      
      stats_list[[m]] <- tibble(
        Metric = m,
        Test = "one-way ANOVA",
        P_value = p_anova
      )
      
      tk <- tryCatch(TukeyHSD(fit), error = function(e) NULL)
      if (!is.null(tk) && "Group" %in% names(tk)) {
        tukey_list[[m]] <- as.data.frame(tk$Group) %>%
          rownames_to_column("Comparison") %>%
          mutate(Metric = m) %>%
          select(Metric, Comparison, everything())
      }
    }
  }
  
  stats_df <- bind_rows(stats_list)
  write.csv(stats_df, file.path(out_dir, "stats_per_metric.csv"), row.names = FALSE)
  
  if (n_groups > 2) {
    tukey_df <- bind_rows(tukey_list)
    if (nrow(tukey_df)) {
      write.csv(tukey_df, file.path(out_dir, "stats_tukeyHSD_all_metrics.csv"), row.names = FALSE)
    }
  }
} else {
  message("Only 1 group selected: skipping statistics (plots will still be generated).")
  write.csv(stats_df, file.path(out_dir, "stats_per_metric.csv"), row.names = FALSE)
}

# ---------------------------- 6) Plotting ------------------------------------

for (m in unique(long_df$Metric)) {
  df_m <- long_df %>% filter(Metric == m)
  if (!nrow(df_m)) next
  
  p <- ggplot(df_m, aes(x = Group, y = Value, fill = Group)) +
    stat_summary(fun = mean, geom = "bar", width = 0.6) +
    stat_summary(fun.data = mean_sd_local, geom = "errorbar", width = 0.2) +
    geom_jitter(width = 0.12, size = 2, alpha = 0.85) +
    labs(title = m, x = NULL, y = m) +
    theme_minimal(base_size = 13) +
    theme(
      legend.position = "none",
      plot.title = element_text(hjust = 0.5)
    )
  
  ggsave(
    filename = file.path(plot_dir, paste0(m, "_barplot.pdf")),
    plot = p,
    width = 6,
    height = 5
  )
}

# ---------------------------- Done --------------------------------------------

message("\nDone ✅ Script 03 completed successfully.")
message("Results written to: ", out_dir)
message("Included groups: ", paste(levels(merged_df$Group), collapse = ", "))

