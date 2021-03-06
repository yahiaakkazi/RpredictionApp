ADDITIONDATAFRAMES_SIMPLE=function(data) {
  x=data.frame(data,Pt=1:length(data[1]),Ecart=1:length(data[1]),Ecart_Absolu=1:length(data[1]))
}

Calcul_Pt_Init_SIMPLE=function(dataset,x) {
  dataset[x,3]=mean(unlist(dataset[2],use.names = FALSE)[1:x])
  return(dataset)
}

Calcul_Pt_SIMPLE=function(dataset,x,alpha) {
  for (i in 1:(length(unlist(dataset[1],use.names = FALSE))-x)) {
    dataset[i+x,3]=alpha*dataset[i+x-1,2]+(1-alpha)*dataset[i+x-1,3]
  }
  return(dataset)
}

Calcul_Ecart_EcartAbsolu_SIMPLE=function(dataset,x) {
  for (i in 1:(length(unlist(dataset[1],use.names = FALSE))-x)) {
    dataset[i+x,4]=dataset[i+x,2]-dataset[i+x,3]
    dataset[i+x,5]=abs(dataset[i+x,4])
  }
  return(dataset)
}

Lissage_Simple=function(Debut,alpha) {
  library("readxl")
  library("writexl")
  x=read_excel(choose.files())
  x=ADDITIONDATAFRAMES_SIMPLE(x)
  y=Calcul_Ecart_EcartAbsolu_SIMPLE(Calcul_Pt_SIMPLE(Calcul_Pt_Init_SIMPLE(x,Debut),Debut,alpha),Debut)
  write_xlsx(y,"C:\\Users\\msi\\Desktop\\R�sultat_Lissage_Simple.xlsx")
  return(y)
}

Visualisation_Graphique_Lissage_Simple=function(x) {
  A=unlist(x[2],use.names = FALSE)
  B=unlist(x[3],use.names = FALSE)
  png(file = "Graphe du lissage simple.jpg")
  plot(A,type = "o",col = "red", xlab = "P�riode", ylab = "Quantit�", main = "Pr�vision sur la demande, M�thode Lissage simple")
  lines(B, type = "o", col = "blue")
  dev.interactive()
  dev.off()
}
require(tcltk)
Fenetre_Simple=function() {
  
  DebutVariable=tclVar("")
  AlphaVariable=tclVar("")
  tt=tktoplevel()
  tkwm.title(tt,"Pr�vision de la demande - M�thode du lissage simple")
  
  Debut.entry=tkentry(tt,textvariable=DebutVariable)
  Alpha.entry=tkentry(tt,textvariable=AlphaVariable)
  
  Lissage_Simplee=function() {
    p=Lissage_Simple(as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable)))
    tkmessageBox(message=paste("Merci de trouver le resultat dans le nouveau excel cr�e dans le bureau sous le nom de R�sultat_Lissage_Simple.xlsx, veuillez aussi cr�er une copie car le document va se faire �craser dans le cas d'un nouveau usage du script"))
    return(p)
  }
  LissageSimple.but=tkbutton(tt, text="M�thode du Lissage Double", command=Lissage_Simplee)
  
  
  Visu_Graphique_Simple=function() {
    Visualisation_Graphique_Lissage_Simple(Lissage_Simple(as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable))))
    tkmessageBox(message=paste("Merci de trouver le graphe dans le dossier 'Documents', veuillez �galement cr�er une copie car l'image va se faire �craser dans le cas d'un nouveau usage du script"))
  }
  VisuLissSimplee.but=tkbutton(tt, text="Visualiser Les graphes Demandes/Pr�visions", command=Visu_Graphique_Simple)
  
  reset=function(){
    tclvalue(DebutVariable)=4
    tclvalue(AlphaVariable)=0.3
  }
  reset.but=tkbutton(tt, text="Remettre les valeurs par d�faut", command=reset)
  tkgrid(tklabel(tt,text="Veuillez choisir la p�riode du d�but"), Debut.entry, pady= 10, padx= 10)
  tkgrid(tklabel(tt,text="Veuillez choisir la valeur d'Alpha"), Alpha.entry, pady= 10, padx= 10)
  tkgrid(reset.but, LissageSimple.but, VisuLissSimplee.but, pady=10, padx=10)
  tkgrid()
  
}
