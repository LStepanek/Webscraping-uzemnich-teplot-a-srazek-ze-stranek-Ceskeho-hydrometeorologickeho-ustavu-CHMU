###############################################################################
###############################################################################
###############################################################################

## websraping stránek ČHMÚ,
## http://portal.chmi.cz/historicka-data/pocasi/uzemni-teploty
## 21. 3. 2017 ----------------------------------------------------------------

###############################################################################

## objektem zájmu je rozsah územních teplot v letech 1961-2016, stejně jako
## územní srážky v leteh 1961-2016 --------------------------------------------

###############################################################################

## vše je dostupné v podobě statického HTML na adresách -----------------------

## http://portal.chmi.cz/files/portal/docs/meteo/ok/uztN_cs.html --------------
## http://portal.chmi.cz/files/portal/docs/meteo/ok/uzsN_cs.html --------------

## např. tedy
## http://portal.chmi.cz/files/portal/docs/meteo/ok/uzt61_cs.html -------------

## kde N je z {61, 62, ..., 99, 00, 01, ..., 16}; obě předchozí stránky se
## skládaj vždy pouze ze statického HTML obsahu -------------------------------

## ----------------------------------------------------------------------------

###############################################################################

## instaluji a loaduji balíčky ------------------------------------------------

for(package in c(
                 "openxlsx"                 
                 )){
                 
    if(!(package %in% rownames(installed.packages()))){
    
        install.packages(
            package,
            dependencies = TRUE,
            repos = "http://cran.us.r-project.org"
        )
        
    }
    
    library(package, character.only = TRUE)
    
}


## ----------------------------------------------------------------------------

###############################################################################

## nastavuji handling se zipováním v R ----------------------------------------

Sys.setenv(R_ZIPCMD = "C:/Rtools/bin/zip") 


## ----------------------------------------------------------------------------

###############################################################################

## nastavuji pracovní složku --------------------------------------------------

while(!"script.R" %in% dir()){
    
    setwd(choose.dir())
    
}

mother_working_directory <- getwd()


## ----------------------------------------------------------------------------

###############################################################################
###############################################################################
###############################################################################

## začínám s webscrapingem ----------------------------------------------------

###############################################################################

## pro každý sufix z {61, 62, ..., 99, 00, 01, ..., 16} stahuji vždy všechna
## statická HTML pro všechny roky odpovídajích sufixům a vždy pro územní
## teploty i srážky -----------------------------------------------------------

###############################################################################

## zakládám pracovní sešit ----------------------------------------------------

weather_conditions <- createWorkbook()


## ----------------------------------------------------------------------------

for(my_condition in c("t", "s")){
    
    my_table <- NULL
    
    
    for(my_suffix in unlist(
        
        lapply(
            c(61:99, 00:05),
            function(x){
                if(nchar(x) == 1){paste("0", x, sep = "")}else{x}
            }
        )
        
    )){
        
        ## načítám HTML obsah pro danou kombinaci podmíny {teplota, srážky}
        ## a daného roku ------------------------------------------------------
    
        my_html <- readLines(
            
            con = paste(
                "http://portal.chmi.cz/files/portal/docs/meteo/ok/uz",
                my_condition,
                my_suffix,
                "_cs.html",
                sep = ""
            ),
            encoding = "latin2"
            
        )
        
        
        ## extrahuji tabulku s hodnotami --------------------------------------
        
        my_rows <- gsub(
            ",",
            ".",
            gsub(
                ";;",
                ";",
                gsub(
                    "<.*?>",
                    ";",
                    gsub(
                        "\t",
                        "",
                        my_html[
                            which(grepl("portlet-table-body", my_html)) +
                            if(
                                all(
                                    nchar(
                                        gsub(
                                            "<.*?>",
                                            "",
                                            my_html[
                                                which(
                                                    grepl(
                                                        "portlet-table-body",
                                                        my_html
                                                    )
                                                )
                                            ]
                                        )
                                    ) == 0
                                )
                            ){1}else{0}
                        ]
                    )
                )
            )
        )
        
        
        ## vymazavám leading a traling semi-colon -----------------------------
        
        my_rows <- unlist(lapply(
            my_rows,
            function(x){
                if(substr(x, 1, 1) == ";"){
                    output <- substr(x, 2, nchar(x))
                }
                if(substr(output, nchar(output), nchar(output)) == ";"){
                    output <- substr(output, 1, nchar(output) - 1)
                }
                return(output)
            }
        ))
        
        my_rows <- lapply(
            lapply(
                my_rows,
                strsplit,
                split = ";"
            ),
            "[[",
            1
        )
        
        
        ## extrahuji pouze číselné hodnoty z tabulky --------------------------
        
        my_values <- lapply(
            my_rows,
            function(x) as.numeric(
                suppressWarnings(x[!is.na(as.numeric(x))])
            )
        )
        
        ## vytvářím tidy data nad hodnotami -----------------------------------
        
        for(i in 1:14){    ## je právě 13 krajů (Praha včetně středních Čech)
                           ## a 14. je celá Česká republika
            
            my_table <- rbind(
                
                my_table,
                cbind(
                    format(
                        my_values[[3 * i - 2]],
                        nsmall = if(my_condition == "t"){1}else{0}
                    ),
                    format(
                        my_values[[3 * i - 1]],
                        nsmall = if(my_condition == "t"){1}else{0}
                    ),
                    format(
                        my_values[[3 * i - 0]],
                        nsmall = if(my_condition == "t"){1}else{0}
                    ),
                    c("leden", "únor", "březen", "duben", "květen",
                      "červen", "červenec", "srpen", "září", "říjen",
                      "listopad", "prosinec", "rok"),
                    rep(
                        unlist(
                            lapply(
                                my_rows,
                                "[[",
                                1
                            )[nchar(lapply(my_rows, "[[", 1)) > 1]
                        )[i],
                        13
                    ),
                    rep(
                        if(as.integer(my_suffix) > 50){
                            paste("19", my_suffix, sep = "")
                        }else{
                            paste("20", my_suffix, sep = "")
                        },
                        13
                    )
                )
                
            )
            
            colnames(my_table) <- c(
                if(my_condition == "t"){
                    "teploty vzduchu [°C]"
                }else{
                    "úhrn srážek [mm]"
                },
                if(my_condition == "t"){
                    "dlouhodobý normál teploty vzduchu 1961-1990 [°C]"
                }else{
                    "dlouhodobý srážkový normál 1961-1990 [mm]"
                },
                if(my_condition == "t"){
                    "odchylka od normálu [°C]"
                }else{
                    "úhrn srážek v % normálu 1961-1990"
                },
                "měsíc",
                "kraj",
                "rok"
            )            
            
        }
        
        flush.console()
        print(
            paste(
                "Stahování dat o územních ",
                if(my_condition == "t"){"teplotách"}else{"srážkách"},
                ": ",
                format(round(which(unlist(
            
                    lapply(
                        c(61:99, 00:05),
                        function(x){
                            if(nchar(x) == 1){paste("0", x, sep = "")}else{x}
                        }
                    )
                    
                ) == my_suffix) / length(unlist(
            
                    lapply(
                        c(61:99, 00:05),
                        function(x){
                            if(nchar(x) == 1){paste("0", x, sep = "")}else{x}
                        }
                    )
                    
                )) * 100, digits = 1), nsmall = 1),
                " %.",
                sep = ""
            )
        )
        
    }
    
    
    ## ukládám výstup do listu sešitu -----------------------------------------

    addWorksheet(
        wb = weather_conditions,
        sheetName = if(my_condition == "t"){"teplota"}else{"srazky"}
    )
    
    writeData(
        wb = weather_conditions,
        sheet = if(my_condition == "t"){"teplota"}else{"srazky"},
        rowNames = FALSE,
        colNames = TRUE,
        x = data.frame(
            my_table,
            check.names = FALSE,
            stringsAsFactors = FALSE
        )
    )
    
    
    ## nastavuji automatickou šířku sloupce -----------------------------------
    
    setColWidths(
        wb = weather_conditions,
        sheet = if(my_condition == "t"){"teplota"}else{"srazky"},
        cols = 1:dim(my_table)[2],
        widths = "auto"
    )
    
    
    ## vytvářím dva své styly - jednak tučné písmo, jednak písmo zarovnané
    ## doprava v rámci buňky --------------------------------------------------

    my_bold_style <- createStyle(textDecoration = "bold")
    right_halign_cells <- createStyle(halign = "right")

    addStyle(
        wb = weather_conditions,
        sheet = if(my_condition == "t"){"teplota"}else{"srazky"},
        style = my_bold_style,
        rows = rep(1, dim(my_table)[2]),
        cols = 1:dim(my_table)[2]
    )

    addStyle(
        wb = weather_conditions,
        sheet = if(my_condition == "t"){"teplota"}else{"srazky"},
        style = right_halign_cells,
        rows = 2:(dim(my_table)[1] + 1),
        cols = 2:(dim(my_table)[2]),
        gridExpand = TRUE
    )
    
}


## ----------------------------------------------------------------------------

###############################################################################
###############################################################################
###############################################################################

## ukládám pracovní sešit -----------------------------------------------------

setwd(mother_working_directory)

if(!file.exists("vystupy")){
    dir.create(file.path(
        mother_working_directory, "vystupy"
    ))
}

setwd(paste(mother_working_directory, "vystupy", sep = "/"))

saveWorkbook(
    wb = weather_conditions,
    file = "uzemni_teploty_a_srazky.xlsx",
    overwrite = TRUE
)

setwd(mother_working_directory)


## ----------------------------------------------------------------------------

###############################################################################
###############################################################################
###############################################################################







