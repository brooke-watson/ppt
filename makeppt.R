makeppt <- function(data, row){ 
  
  #install required libraries
  library(dplyr)
  library(magrittr)
  library(tidyr)
  library(ggplot2)
  library(ReporteRs)
  library(foreign)
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
