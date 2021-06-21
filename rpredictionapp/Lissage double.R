ADDITIONDATAFRAMES_DOUBLE=function(data) {
  x=data.frame(data,Lt=1:length(data[1]),Tt=1:length(data[1]),Pt=1:length(data[1]),Ecart=1:length(data[1]),Ecart_Absolu=1:length(data[1]))
}

Calcul_Lt_Tt_Pt_Init_DOUBLE=function(dataset,x) {
  dataset[x,3]=mean(unlist(dataset[2],use.names = FALSE)[1:x])
  dataset[x,4]=(dataset[x,2]-dataset[1,2])/(dataset[x,1]-dataset[1,1])
  dataset[x,5]=dataset[x,3]+dataset[x,4]
  return(dataset)
}

Calcul_Lt_Tt_Pt_DOUBLE=function(dataset,x,alpha,beta) {
  for (i in 1:(length(unlist(dataset[1],use.names = FALSE))-x)) {
    dataset[i+x,3]=alpha*dataset[i+x-1,2]+(1-alpha)*dataset[i+x-1,5]
    dataset[i+x,4]=beta*(dataset[i+x,3]-dataset[i+x-1,3])+(1-beta)*dataset[i+x-1,4]
    dataset[i+x,5]=dataset[i+x,4]+dataset[i+x,3]
  }
  return(dataset)
}

Calcul_Ecart_EcartAbsolu_DOUBLE=function(dataset,x) {
  for (i in 1:(length(unlist(dataset[1],use.names = FALSE))-x)) {
    dataset[i+x,6]=dataset[i+x,2]-dataset[i+x,5]
    dataset[i+x,7]=abs(dataset[i+x,6])
  }
  return(dataset)
}

Lissage_Double=function(Debut,alpha,beta) {
  library("readxl")
  library("writexl")
  x=read_excel(choose.files())
  x=ADDITIONDATAFRAMES_DOUBLE(x)
  y=Calcul_Ecart_EcartAbsolu_DOUBLE(Calcul_Lt_Tt_Pt_DOUBLE(Calcul_Lt_Tt_Pt_Init_DOUBLE(x,Debut),Debut,alpha,beta),Debut)
  write_xlsx(y,"C:\\Users\\msi\\Desktop\\Résultat_Lissage_Double.xlsx")
  return(y)
}

Visualisation_Graphique_Lissage_Double=function(x) {
  A=unlist(x[2],use.names = FALSE)
  B=unlist(x[5],use.names = FALSE)
  png(file = "Graphe du lissage double.jpg")
  plot(A,type = "o",col = "red", xlab = "Période", ylab = "Quantité", main = "Prévision sur la demande, Méthode Lissage double")
  lines(B, type = "o", col = "blue")
  dev.interactive()
  dev.off()
}

require(tcltk)
Fenetre_Double=function() {
  
  DebutVariable=tclVar("")
  AlphaVariable=tclVar("")
  BetaVariable=tclVar("")
  tt=tktoplevel()
  tkwm.title(tt,"Prévision de la demande - Différents Méthodes")
  
  Debut.entry=tkentry(tt,textvariable=DebutVariable)
  Alpha.entry=tkentry(tt,textvariable=AlphaVariable)
  Beta.entry=tkentry(tt,textvariable=BetaVariable)
  
  Lissage_Doublee=function() {
    p=Lissage_Double(as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable)),as.double(tclvalue(BetaVariable)))
    tkmessageBox(message=paste("Merci de trouver le resultat dans le nouveau excel crée dans le bureau sous le nom de Résultat_Lissage_Double.xlsx, veuillez aussi créer une copie car le document va se faire écraser dans le cas d'un nouveau usage du script"))
    return(p)
  }
  LissageDouble.but=tkbutton(tt, text="Méthode du Lissage Double", command=Lissage_Doublee)

  
  Visu_Graphique_Double=function() {
    Visualisation_Graphique_Lissage_Double(Lissage_Double(as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable)),as.double(tclvalue(BetaVariable))))
    tkmessageBox(message=paste("Merci de trouver le graphe dans le dossier 'Documents', veuillez également créer une copie car l'image va se faire écraser dans le cas d'un nouveau usage du script"))
  }
  VisuLissDouble.but=tkbutton(tt, text="Visualiser Les graphes Demandes/Prévisions", command=Visu_Graphique_Double)
  
  reset=function(){
    tclvalue(DebutVariable)=4
    tclvalue(AlphaVariable)=0.3
    tclvalue(BetaVariable)=0.2
  }
  reset.but=tkbutton(tt, text="Remettre les valeurs par défaut", command=reset)
  tkgrid(tklabel(tt,text="Veuillez choisir la période du début"), Debut.entry, pady= 10, padx= 10)
  tkgrid(tklabel(tt,text="Veuillez choisir la valeur d'Alpha"), Alpha.entry, pady= 10, padx= 10)
  tkgrid(tklabel(tt,text="Veuillez choisir la valeur de Beta"), Beta.entry, pady= 10, padx= 10)
  tkgrid(reset.but, LissageDouble.but,VisuLissDouble.but, pady=10, padx=10)
  tkgrid()
  
}