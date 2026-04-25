# rk.exporter: Batch Plot and Table Exporting for RKWard

![Version](https://img.shields.io/badge/Version-0.0.1-blue.svg)
![License](https://img.shields.io/badge/License-GPLv3-blue.svg)
![RKWard](https://img.shields.io/badge/Platform-RKWard-green)
[![R Linter](https://github.com/AlfCano/rk.exporter/actions/workflows/lintr.yml/badge.svg)](https://github.com/AlfCano/rk.exporter/actions/workflows/lintr.yml)

**rk.exporter** brings seamless, automated batch exporting capabilities to the RKWard GUI. Stop wasting hours manually exporting plots one by one or copy-pasting tables into Microsoft Office. 

Powered by `purrr`, `flextable`, and `officer`, this plugin suite allows you to take a List of `ggplot2` objects or `flextables` and instantly export them as hundreds of individual high-res files, or compile them directly into professional, multi-page Word documents, PDFs, or PowerPoint presentations.

## 🚀 What's New in Version 0.0.1

This is the initial release of the package, introducing two core automation tools for data reporting:

1.  **Batch Plot Exporter:** Export lists of `ggplot2` graphics. Choose between saving them as individual files (SVG, PNG, PDF) into a folder, or combining them into a single presentation/report (PPTX, DOCX, PDF).
2.  **Batch Table Exporter:** Export lists of `flextable` objects. Output them as individual HTML/Word/PPTX files, or stack them perfectly into a single, paginated Word document or PowerPoint deck.

### 🌍 Internationalization
The interface is fully localized in:
*   🇺🇸 English (Default)
*   🇪🇸 Spanish (`es`)
*   🇫🇷 French (`fr`)
*   🇩🇪 German (`de`)
*   🇧🇷 Portuguese (Brazil) (`pt_BR`)

## ✨ Features

### 1. Smart Combined Documents (Office Integration)
*   **One Object per Page:** When combining tables or plots into Word (`.docx`), the plugin automatically injects page breaks (`body_add_break()`) so every table/plot sits cleanly on its own page.
*   **Orientation Control:** Instantly set the entire Word or PDF document to **Landscape** or **Portrait** mode.
*   **PowerPoint Titles:** When exporting a named list to PPTX, the name of the list element automatically becomes the title of the slide!

### 2. Bulletproof Batch Processing
*   **Name Sanitization:** Safely handles list names containing spaces, slashes, or special characters, converting them to safe filenames (e.g., `My Plot / 2` -> `My_Plot___2.png`) to prevent OS errors.
*   **Auto-Naming:** If you pass an unnamed list (e.g., from an `lapply` loop), the plugin will automatically generate sequential names (`plot_1`, `plot_2`, etc.).
*   **Auto-Extensions:** Don't worry about typing the exact file extension; the plugin's logic checks and appends `.docx` or `.pptx` dynamically if you forget it.

### 3. Dynamic User Interface
*   **Error Prevention:** The GUI uses dynamic logic (`rk.XML.logic`). Depending on whether you choose "Individual" or "Combined" modes, unnecessary fields grey out and become disabled, guiding you to a perfect export every time.

## 📦 Installation

This plugin is not yet on CRAN. To install it, use the `remotes` or `devtools` package in RKWard.

1.  **Open RKWard**.
2.  **Run the following command** in the R Console:

    ```R
    # If you don't have devtools installed:
    # install.packages("devtools")
    
    local({
      require(devtools)
      install_github("AlfCano/rk.exporter", force = TRUE)
    })
    ```
3.  **Restart RKWard** to load the new menu entries.

## 💻 Usage & Testing Instructions

Once installed, the tools are located under the File menu:
**`File` -> `Batch Exporters`**

To test the suite, copy and paste this code into your RKWard console to instantly generate a list of ggplot2 graphics and a list of flextables:

```R
library(ggplot2)
library(flextable)

# 1. Create a named list of ggplot2 objects
p1 <- ggplot(mtcars, aes(x = wt, y = mpg, color = factor(cyl))) + geom_point(size=3) + theme_minimal() + labs(title="MTCars Scatter")
p2 <- ggplot(iris, aes(x = Species, y = Sepal.Length, fill=Species)) + geom_boxplot() + theme_minimal() + labs(title="Iris Boxplot")
my_plots <- list("Car_Analysis" = p1, "Flower_Analysis" = p2)

# 2. Create a named list of flextable objects
t1 <- flextable(head(mtcars[, 1:5])) %>% theme_vanilla() %>% set_caption("Motor Trend Cars")
t2 <- flextable(head(iris)) %>% theme_zebra() %>% set_caption("Edgar Anderson's Iris Data")
my_tables <- list("Cars_Table" = t1, "Iris_Table" = t2)
```

### Test 1: Export Plots to PowerPoint
1. Go to **File -> Batch Exporters -> Batch Plot Exporter**.
2. **Input & Mode:** Select `my_plots` as your Target object.
3. Choose **Combined Document**. Notice how the "Output Directory" greys out!
4. Choose an **Output File** (e.g., your Desktop, name it `Report`).
5. **Export Settings:** Select **PowerPoint (.pptx)**. 
6. Click **Submit**. Open the PPTX file: You will have two slides, each titled automatically ("Car_Analysis" and "Flower_Analysis") with high-quality vector plots inside!

### Test 2: Export Tables to Individual Word Files
1. Go to **File -> Batch Exporters -> Batch Table Exporter**.
2. **Input & Mode:** Select `my_tables`.
3. Choose **Individual Files**. Select an Output Directory (e.g., a new folder on your Desktop).
4. **Export Settings:** Select **Word (.docx)** and **Landscape** orientation.
5. Click **Submit**. Check your folder: You will see `Cars_Table.docx` and `Iris_Table.docx` perfectly exported in landscape mode.

## 🛠️ Dependencies

This plugin relies on the following R packages:
*   `purrr` (List iteration and mapping)
*   `ggplot2` (Graphics rendering and `ggsave`)
*   `svglite` (High-quality SVG device rendering)
*   `flextable` (Table formatting)
*   `officer` (Native Microsoft Office document manipulation)
*   `rkwarddev` (Plugin generation)

## ✍️ Author & License

*   **Author:** Alfonso Cano (<alfonso.cano@correo.buap.mx>)
*   **Assisted by:** Gemini, a large language model from Google.
*   **License:** GPL (>= 3)
