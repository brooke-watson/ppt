#Create slides for each of the countries.

#setup 
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
wd <- "~/Dropbox (EHA)/GVP/global-virome-master/powerpoint"
setwd(wd)

#loading in the data 
db <- read.csv("~/Dropbox (EHA)/GVP/global-virome-master/powerpoint/Mammals_per_country.csv")
View(db)
head(db)
str(db)

#Example graphs using values for India. 
#Eventually these values will be auto-populated from a database. 
countries <- db$country
#country <- as.character(db$country) # careful with this. leave as India for now. 
country <- "India"
nmam <- db$nmam
nbirds <- db$nbirds
biodivreg <- db$biodivreg
hsreg <- db$hsreg
nzoo <- 500

rm(biodivregion)
#first bullet point on the mammal biodiversity slide - some have data for waterfowl while 
#others only have data for waterbirds (broader). Distinguishes. 
# if (is.na(db$nfowl) == FALSE){ 
#     mb1 <- paste0("~", paste0(nmam)," Mammal species reside in ", paste0(country), "; ~", paste0(nfowl)," species of water fowl")
#         elseif (is.na(db$nfowl) == TRUE & is.na(db$nbirds) == FALSE){
#         mb1 <- paste0("~", paste0(nmam)," Mammal species reside in ", paste0(country), "; ~", paste0(nbirds)," species of water birds")
#         }
#             else {
#             mb1 <- paste0("~", paste0(nmam)," Mammal species reside in ", paste0(country))    
#             }
# }        

mb1 <- paste0("~", paste0(nmam)," Mammal species reside in ", paste0(country), "; ~", paste0(nbirds)," species of water birds")

mb1 <- sprintf("%s Mammal species reside in %s; along with %s species of waterbirds", nmam, country, nbirds)
mb2 <- "Many are migratory or occur in other countries"
mb3 <- sprintf("Densest areas of biodiversity located in the %s", biodivreg)


country <- data[row, 1]
hsreg <- data[row, "hsreg"] 
nmam <- data[row, ]

#bullet point template for each slide  
mambullets <- c(mb1, "", mb2, "", mb3)  

hsbullets <- c("Mammal biodiversity, population density, and land use patterns are key drivers in emergence of infectious diseases",
    sprintf("Risk is elevated across much of %s, particularly in the %s", country, hsreg))

 
zoobullets <- c("This identifies hotspots where we would find the most new zoonotic viruses if we sampled uniformly all mammals",
sprintf("Over %s unidentified wildlife zoonoses predicted in %s based on modelled analysis (Olival et al. in review, Nature)", nzoo, country),
"Only a subset of mammals studied – actual number of viruses could be higher")

#makeslide <- function(country){}     # leave for now 

row.names(db)

getrow <- function(data, n) {
    #grab just the row you want 
    data<- data[n,]
    return(data)
}
getname <- function(row){
    name <- as.character(row[,1])
    return(name)
}

afg <-  getrow(db, 1)
 
 
makeppt(data, row){ 
    #select the country 
    name <- as.character(paste0(getname(getrow(data, row))) 
                         
    #                     
                         
                         
                         
                         
    doc = pptx(sprintf("%s GVP test.pptx", name), template = "GVPmaptemp.pptx")
doc <- doc %>% 
    addSlide( slide.layout = "Custom_for_Maps" ) %>% 
        addTitle(paste0(name, "- Mammal Biodiversity")) %>% 
        addParagraph((value = mambullets)) %>% 
        addPlot( function( ) print( plot1 ) ) %>%  
    addSlide(slide.layout = "Custom_for_Maps" ) %>% 
        addTitle(paste0(name, "- EID Hotspots")) %>% 
        addParagraph((value = mambullets)) %>% 
        addPlot( function( ) print( plot1 ) ) %>%  
    addSlide( slide.layout = "Custom_for_Maps" ) %>% 
        addTitle(paste0(name, " - Missing Zoonoses")) %>% 
        addParagraph((value = zoobullets)) %>% 
        addPlot( function( ) print( plot1 ) ) %>% 
        writeDoc(paste0(country, " GVP map.pptx"))   
}
    
#testing   
# writeDoc(paste0(getname(getrow(data, row)), " GVP map.pptx"))
# writeDoc(doc, paste0(getname(getrow(db, 5)), " GVP map.pptx"))
# writeDoc(doc, paste0(name), " GVP map.pptx"))

#### References

```{r bibtext, echo = FALSE, message = FALSE, cache=FALSE, results = 'asis'}
write.bibtex(file="references.bib")
```

