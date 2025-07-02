# ─────────────────────────────────────────────────────────────────────────────
# SETUP & LIBRARIES
# ─────────────────────────────────────────────────────────────────────────────

library(readr)      # read_csv()
library(dplyr)      # data manipulation
library(tidyr)      # complete(), fill()
library(lubridate)  # make_date(), %m+%
library(stringr)    # str_to_title()
library(forecast)   # auto.arima(), forecast()
library(purrr)      # map(), map2()
library(tibble)     # tibble()

# ─────────────────────────────────────────────────────────────────────────────
# 1) READ & CLEAN
# ─────────────────────────────────────────────────────────────────────────────

df <- read_csv("Case Study.csv") %>%
  mutate(
    Month = str_to_title(Month),                                       # normalize month names
    Date  = make_date(Year, match(Month, month.name), 1)               # build Date
  ) %>%
  group_by(Item, Customer) %>%
  mutate(
    # backfill any missing statuses from within each Item×Customer
    `Item Status` = coalesce(`Item Status`, first(na.omit(`Item Status`)))
  ) %>%
  ungroup()

# ─────────────────────────────────────────────────────────────────────────────
# 2) FILTER OUT DISCONTINUED (keep Active, Watch List, and NAs)
# ─────────────────────────────────────────────────────────────────────────────

df_ship <- df %>%
  filter(`Item Status` != "Discontinued") %>%   # removes only Discontinued
  select(Customer, Item, Description, `Item Status`, Date, Units)

# ─────────────────────────────────────────────────────────────────────────────
# 3) BUILD FULL GRID & ZERO-FILL
# ─────────────────────────────────────────────────────────────────────────────

# determine historical date bounds
date_bounds <- df_ship %>%
  group_by(Item, Customer) %>%
  summarize(
    first_date = min(Date),
    last_date  = max(Date),
    .groups    = "drop"
  )

# expand to every month in between
full_grid <- date_bounds %>%
  mutate(Date = map2(first_date, last_date, ~ seq(.x, .y, by = "month"))) %>%
  select(-first_date, -last_date) %>%
  unnest(Date)

# join back and fill
df_ship_filled <- full_grid %>%
  left_join(df_ship, by = c("Item","Customer","Date")) %>%
  arrange(Item, Customer, Date) %>%
  group_by(Item, Customer) %>%
  fill(Description, `Item Status`, .direction = "downup") %>%  # carry known values up/down
  ungroup() %>%
  mutate(Units = replace_na(Units, 0))

# ─────────────────────────────────────────────────────────────────────────────
# 4) OUTLIER DETECTION (1.5×IQR per Item)
# ─────────────────────────────────────────────────────────────────────────────
# 1) Compute per-series IQR bounds
iqr_bounds <- df_ship_filled %>%
  group_by(Item, Customer) %>%
  summarize(
    Q1   = quantile(Units, .25, na.rm = TRUE),
    Q3   = quantile(Units, .75, na.rm = TRUE),
    IQR  = Q3 - Q1,
    low  = Q1 - 1.5 * IQR,
    high = Q3 + 1.5 * IQR,
    .groups = "drop"
  )

# 2) Join back on BOTH Item & Customer, then flag outliers
outliers <- df_ship_filled %>%
  left_join(iqr_bounds, by = c("Item", "Customer")) %>%
  mutate(
    is_outlier   = Units < low | Units > high,
    outlier_type = case_when(
      Units < low  ~ "Below Lower Bound",
      Units > high ~ "Above Upper Bound",
      TRUE         ~ NA_character_
    )
  ) %>%
  filter(is_outlier) %>%
  select(Item, Customer, Date, Units, outlier_type)

# 3) Inspect
print(outliers)


# ─────────────────────────────────────────────────────────────────────────────
# 5) FORECASTING (6‐Month SARIMA per Item×Customer)
# ─────────────────────────────────────────────────────────────────────────────

ts_tbl <- df_ship_filled %>%
  group_by(Item, Customer) %>%
  arrange(Date) %>%
  summarize(
    series     = list(Units),
    start_date = first(Date),
    .groups    = "drop"
  )

ts_fc <- ts_tbl %>%
  mutate(
    ts  = map2(series, start_date, ~ ts(.x,
                                        start     = c(year(.y), month(.y)),
                                        frequency = 12)),
    fit = map(ts, auto.arima,
              seasonal     = TRUE,
              stepwise     = FALSE,
              approximation= FALSE),
    fc  = map(fit, forecast, h = 6)
  )

# … assume everything up through ts_fc is as before …

last_date <- max(df_ship_filled$Date)

forecasts <- ts_fc %>%
  transmute(
    Item,
    Customer,
    forecast = map(fc, ~ tibble(
      Date       = seq(last_date %m+% months(1), by = "month", length.out = 6),
      Mean       = as.numeric(.x$mean),
      Lo80       = .x$lower[,1],
      Hi80       = .x$upper[,1],
      Lo95       = .x$lower[,2],
      Hi95       = .x$upper[,2]
    ))
  ) %>%
  unnest(forecast) %>%
  mutate(
    Rounded = ceiling(Mean)   # new column, rounded up to nearest integer
  )

print(forecasts)

#make fcst wide




# ─────────────────────────────────────────────────────────────────────────────
# WRITE OUTLIERS & FORECASTS TO EXCEL
# ─────────────────────────────────────────────────────────────────────────────

# install.packages("writexl")   # if you don’t already have it
library(writexl)

# Make a named list of your two data frames
sheets <- list(
  Outliers  = outliers,
  Forecasts = forecasts
)

# Write them to “Summary.xlsx” in your working directory
write_xlsx(sheets, path = "Summary.xlsx")

message("Wrote Summary.xlsx with sheets: ", paste(names(sheets), collapse = ", "))

