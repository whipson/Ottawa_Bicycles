library(tidyverse)
library(readxl)
library(lubridate)
library(devtools)
library(janitor)

read_excel_allsheets <- function(filename, tibble = TRUE) {
  sheets <- readxl::excel_sheets(filename)
  x <- lapply(sheets, function(X) readxl::read_excel(filename, sheet = X))
  if(!tibble) x <- lapply(x, as.data.frame)
  names(x) <- paste0("bikes_", sheets)
  x
}

bikes <- read_excel_allsheets("bike_counter 19-10-23.xlsx")

list2env(bikes, envir = .GlobalEnv)

bikes_2010 <- bikes_2010 %>%
  slice(-1) %>%
  mutate(Date = parse_date_time(Date, orders = "dmy")) %>%
  mutate_at(c(-1), as.numeric) %>%
  rename(`1^ALEX` = `1_ALEX`,
         `2^ORPY` = `2_ORPY`,
         `3^COBY` = `3_COBY`)

bikes_2011 <- bikes_2011 %>%
  slice(-1) %>%
  mutate(Date = parse_date_time(Date, orders = "dmy")) %>%
  mutate_at(c(-1), as.numeric) %>%
  rename(`1^ALEX` = AlexS2,
         `2^ORPY` = ORP_PoW3,
         `3^COBY` = ColByEL2,
         `4^CRTZ` = `CNL RTZ_6892`,
         `5^LMET` = 8,
         `6^LLYN` = 7,
         `7^LBAY` = 6)

bikes_2012 <- bikes_2012 %>%
  slice(-1) %>%
  mutate(date1 = parse_date_time(Date, orders = "dmy"),
         date2 = as.POSIXct(as.Date(as.numeric(Date), origin = "1899-12-30"))) %>%
  mutate(Date = coalesce(date1, date2)) %>%
  select(-date1, -date2) %>%
  mutate_at(c(-1), as.numeric) %>%
  rename(`1^ALEX` = AlexS2,
         `2^ORPY` = ORP_PoW3,
         `3^COBY` = ColByEL2,
         `4^CRTZ` = `CNL RTZ_6892`,
         `5^LMET` = Laurier_at_Metcalfe,
         `6^LLYN` = `LSBL lyo`,
         `7^LBAY` = `LSBL Bay`,
         `8^SOMO` = Sommerset_OTrain_OUT)

bikes_2013 <- bikes_2013 %>%
  slice(1:(nrow(.)-2)) %>%
  mutate(Date = as.POSIXct(as.Date(as.numeric(Date), origin = "1899-12-30"))) %>%
  mutate_at(c(-1), as.numeric) %>%
  rename(`9 OYNG 1` = `9 OYNG`,
         `11 OBVW` = `11^OBVW`)

bikes_2014 <- bikes_2014 %>%
  mutate_at(c(-1), as.numeric)

bikes_2016 <- bikes_2016 %>%
  mutate(date1 = parse_date_time(Date, orders = "dmy"),
         date2 = as.POSIXct(as.Date(as.numeric(Date), origin = "1899-12-30"))) %>%
  mutate(Date = coalesce(date1, date2)) %>%
  select(-date1, -date2)

bikes_2017 <- bikes_2017 %>%
  mutate(Date = parse_date_time(Date, orders = "dmy"))

bikes_2019 <- bikes_2019 %>%
  slice(1:273) %>%
  mutate(Date = as.POSIXct(as.Date(as.numeric(Date), origin = "1899-12-30")),
         `1^ALEX` = as.numeric(`1^ALEX`))

bikes_all <- bind_rows(bikes_2010, bikes_2011, bikes_2012, bikes_2013, bikes_2014, bikes_2015,
                       bikes_2016, bikes_2017, bikes_2018, bikes_2019) %>%
  clean_names() %>%
  rename(alexandra_bridge = x1_alex,
         ottawa_river = x2_orpy,
         eastern_canal = x3_coby,
         western_canal = x4_crtz,
         laurier_metcalfe = x5_lmet,
         laurier_lyon = x6_llyn,
         laurier_bay = x7_lbay,
         somerset_bridge = x8_somo,
         otrain_young = x9_oyng_1,
         otrain_gladstone = x10_ogld,
         otrain_bayview = x11_obvw,
         portage_bridge = portage_bridge,
         adawe_crossing_a = x12a_adawe,
         adawe_crossing_b = x12b_adawe)

write_csv(bikes_all, "ottawa_bike_counts.csv")


