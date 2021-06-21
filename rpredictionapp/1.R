y=function() {
  x=as.double(readline(prompt = "Nombre de périodes de visualisation ?"))
  y=1:x
  return(y)
}
S=data.frame(y())library("readxl")z