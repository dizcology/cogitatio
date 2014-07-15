fton <- function(f){
  return(as.numeric(levels(f))[f])
}

percent <- function(v,th){
  return(length(which(v>=th))/length(v)*100)
}

