# #Create slides for each of the countries.
# pkgs <- c("dplyr", "magrittr", "tidyr", "ggplot2", "viridis", "deSolve", "knitcitations", 
#          "purrr", "animation", "gganimate", "ReporteRs", "foreign")
# 
# install.packages(pkgs)
# #setup 
library(dplyr)
library(magrittr)
library(tidyr)
library(ggplot2)
library(viridis)
library(deSolve)
library(knitcitations)
library(purrr)
library(animation)
library(gganimate)
library(ReporteRs)
library(foreign)

#cleanup 
rm(list=ls())
# wd <- "~/Dropbox (EHA)/GVP/global-virome-master/powerpoint"
# setwd(wd)

#loading in the data 
#db <- read.csv("~/Dropbox (EHA)/GVP/global-virome-master/powerpoint/Mammals_per_country.csv")
db <- read.csv("~/Desktop/git/ppt/Mammals_per_country.csv")

# 
# View(db)
# head(db)
# str(db)

#Example graphs using values for India. 
#Eventually these values will be auto-populated from a database. 
#this whole thing maybe not useful: 

# countries <- db$country
# #country <- as.character(db$country) # careful with this. leave as India for now. 
# country <- "India"
# nmam <- db$nmam
# nbirds <- db$nbirds
# biodivreg <- db$biodivreg
# hsreg <- db$hsreg
# nzoo <- 500

# first bullet point on the mammal biodiversity slide - some have data for waterfowl while
# others only have data for waterbirds (broader). Distinguishes.
# if (is.na(db$nfowl) == FALSE){
#     mb1 <- paste0("~", paste0(nmam)," Mammal species reside in ", paste0(country), "; ~", paste0(nfowl)," species of water fowl")
#         elseif (is.na(db$nfowl) == TRUE & is.na(db$nbirds) == FALSE){
#         mb1 <- paste0("~", paste0(nmam)," Mammal species reside in ", paste0(country), "; ~", paste0(nbirds)," species of water birds")
#         }
#             else {
#             mb1 <- paste0("~", paste0(nmam)," Mammal species reside in ", paste0(country))
#             }
# }

#use sprintf instead 
# mb1 <- paste0("~", paste0(nmam)," Mammal species reside in ", paste0(country), "; ~", paste0(nbirds)," species of water birds")

 

#bullet point template for each slide  


 

#makeslide <- function(country){}     # leave for now 


#this is a mess. just use indices: 
# getrow <- function(data, n) {
#     #grab just the row you want 
#     data<- data[n,]
#     return(data)
# }
# getname <- function(row){
#     name <- as.character(row[,1])
#     return(name)
# }

# afg <-  getrow(db, 1)
 
# getlist <- function(data, row, varlist)  
#     for(i in vars) {
#     paste(i) <- data[row, i] 
#     return(paste(i))  
#     }
# getlist(db, 3, vars)
#     
#     #example 
#     foreach (i in vars) {
#       i <- data[1, as.character(i)] 
#       
#     }
# this doesn't work. 

makeppt <- function(data, row){ 
    #select the country 
    vars <- c("country", "nmam", "nbirds", "hsreg", "biodivreg", "nzoo")  
    # easier: vars <- colnames(db)

      country <- data[row, "country"]
      nmam <- data[row, "nmam"]
      nbirds <- data[row, "nbirds"]
      hsreg <- data[row, "hsreg"] 
      biodivreg <- data[row, "biodivreg"]
      nzoo <- data[row, "nzoo"]
    
    #define the values in the bullet points that will vary by country 
    

#############################     MAKE BULLET POINTS     #############################
    
    #make the bullets for the first slide 
    mb1 <- sprintf("%s Mammal species reside in %s; along with %s species of waterbirds", nmam, country, nbirds)
    mb2 <- "Many are migratory or occur in other countries"
    mb3 <- sprintf("Densest areas of biodiversity located in the %s", biodivreg)
    mambullets <- c(mb1, "", mb2, "", mb3) 
    
    #make the bullets for the second slide 
    hsbullets <- c("Mammal biodiversity, population density, and land use patterns are key drivers in emergence of infectious diseases","",
        sprintf("Risk is elevated across much of %s, particularly in the %s", country, hsreg))
    
    #make the bullets for the third slide 
    zoobullets <- c("This identifies hotspots where we would find the most new zoonotic viruses if we sampled uniformly all mammals","",
                    sprintf("Over %s unidentified wildlife zoonoses predicted in %s based on modelled analysis (Olival et al. in review, Nature)", nzoo, country),
                    "", "Only a subset of mammals studied – actual number of viruses could be higher")
    #define the values before 
    
    doc = pptx(sprintf("%s GVP test.pptx", country), template = "GVPmaptemp.pptx")
doc <- doc %>% 
    addSlide( slide.layout = "Custom_for_Maps" ) %>% 
        addTitle(paste0(country, "- Mammal Biodiversity")) %>% 
        addParagraph((value = mambullets )) %>% 
        addImage("Mammal_biodiversity_FINAL.jpg", width=4.4, height=5.2) %>% 
  addSlide(slide.layout = "Custom_for_Maps" ) %>% 
        addTitle(paste0(country, "- EID Hotspots")) %>% 
        addParagraph((value = hsbullets)) %>% 
        addImage("India_hotspots1-2.5_FINAL.jpg", width=4.4, height=5.2) %>% 
    addSlide( slide.layout = "Custom_for_Maps" ) %>% 
        addTitle(paste0(country, " - Missing Zoonoses")) %>% 
        addParagraph((value = zoobullets)) %>% 
        addImage("Missing_Zoonoses_FINAL.jpg", width=4.4, height=5.2) %>% 
    writeDoc(paste0(country, " GVP map.pptx"))   
}

makeppt(db, 98)    
#testing   
# writeDoc(paste0(getname(getrow(data, row)), " GVP map.pptx"))
# writeDoc(doc, paste0(getname(getrow(db, 5)), " GVP map.pptx"))
# writeDoc(doc, paste0(name), " GVP map.pptx"))

#### References

```{r bibtext, echo = FALSE, message = FALSE, cache=FALSE, results = 'asis'}
write.bibtex(file="references.bib")
```

