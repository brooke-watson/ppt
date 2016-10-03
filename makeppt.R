######################################################################################### 
####                    GVP PPTs
######################################################################################### 

 
db <- read.csv("~/Dropbox (EHA)/gvp (1)/notes/GVP mammal species.csv", stringsAsFactors = FALSE)
 
makeppt <- function(data, row){ 
  
  #install required libraries
  library(dplyr)
  library(magrittr)
  library(tidyr)
  library(ggplot2)
  library(ReporteRs)
  library(foreign)
  #select the country 
     
  country <- data[row, 1]
  nmam <- data[row, 3]
  # nbirds <- data[row, "nbirds"]
 
  
  #define the values in the bullet points that will vary by country 
  
  
  #############################     MAKE BULLET POINTS     #############################
  
  #make the bullets for the first slide 
  mb1 <- sprintf("%s Mammal species reside in %s", nmam, country)
  mb2 <- "Many species also occur in other countries"
  mb3 <- sprintf("Densest areas of mammal biodiversity are located in the red areas, but some areas of low biodiversity are highly abundant")
  mambullets <- c(mb1, "", mb2, "", mb3) 
  
  birdbullets <- c("~40 species of Anatidae spp reside throughout India, of roughly 140 species worldwide", "", 
                    "Migratory birds in the Anatidae family have been implicated in the spread of pandemic influenza")
  
  #make the bullets for the second slide 
  hsbullets <- c("Mammal biodiversity, population density, and land use patterns are key drivers in emergence of infectious diseases","",
                 sprintf("Risk is elevated across much of %s, particularly in the red regions", country))
  
  #make the bullets for the third slide 
  zoobullets <- c("This identifies hotspots where we would find the most new zoonotic viruses if we sampled uniformly all mammals","",
                  sprintf("Over 200 unidentified wildlife zoonoses predicted in %s based on modelled analysis (Olival et al. in review, Nature)", country)) #,
                  #"", "Only a subset of mammals studied – actual number of viruses could be higher")
  #last bullet commented out because it was too long. 
  
  doc = pptx(sprintf("%s GVP maps.pptx", country), template = "GVPmaptemp.pptx")
  doc <- doc %>% 
    addSlide( slide.layout = "Custom_for_Maps" ) %>% 
    addTitle(paste0(country, "- Mammal Biodiversity")) %>% 
    addParagraph((value = mambullets )) %>% 
 #   addImage(paste0(country, "- Mammal Biodiversity.jpg"), width=4.4, height=5.2) %>% 
    addSlide( slide.layout = "Custom_for_Maps" ) %>% 
    addTitle(paste0(country, "- Anatidae Biodiversity")) %>% 
    addParagraph((value = birdbullets )) %>% 
 #   addImage(paste0(country, "- Anatidae Biodiversity.jpg"), width=4.4, height=5.2) %>% 
    addSlide(slide.layout = "Custom_for_Maps" ) %>% 
    addTitle(paste0(country, "- EID Hotspots")) %>% 
    addParagraph((value = hsbullets)) %>% 
 #   addImage(paste0(country, "- EID Hotspots.jpg"), width=4.4, height=5.2) %>% 
    addSlide( slide.layout = "Custom_for_Maps" ) %>% 
    addTitle(paste0(country, " - Missing Zoonoses")) %>% 
    addParagraph((value = zoobullets)) %>% 
 #   addImage(paste0(country, "Missing_Zoonoses.jpg"), width=4.4, height=5.2) %>% 
    writeDoc(paste0(country, " GVP map.pptx"))   
}

lapply(db[1:34,], makeppt, db) 

for (n in 6:34) {

makeppt(db, n)

}
lapply(c(3:5), makeppt, db)