
ADDITIONDATAFRAMES=function(data) {
  x=data.frame(data,St=1:length(data[1]),Lt=1:length(data[1]),Tt=1:length(data[1]),Pt=1:length(data[1]),Pt_corrigée=1:length(data[1]),Ecart=1:length(data[1]),Ecart_Absolu=1:length(data[1]))
  
}
Calcul_St_Lt_Tt_Pt_Init=function(dataset,x) {
  for (i in 1:x)
    dataset[i,3]=dataset[i,2]/mean(unlist(dataset[2],use.names = FALSE)[1:x])
  dataset[x,4]=mean(unlist(dataset[2],use.names = FALSE)[1:x])
  dataset[x,5]=(dataset[x,2]-dataset[1,2])/(dataset[x,1]-dataset[1,1])
  dataset[x,6]=dataset[x,5]+dataset[x,4]
  return(dataset)
}

Calcul_St_Lt_Tt_Pt=function(dataset,x,alpha,beta,gamma) {
  for (i in 1:(length(unlist(dataset[1],use.names = FALSE))-x)) {
    dataset[i+x,3]=gamma*(dataset[i+x-1,2]/dataset[i+x-1,4])+(1-gamma)*dataset[i,3]
    dataset[i+x,4]=alpha*(dataset[i+x-1,2]/dataset[i,3])+(1-alpha)*dataset[i+x-1,6]
    a=dataset[i+x,4]
    dataset[i+x,5]=beta*(a-dataset[i+x-1,4])+(1-beta)*dataset[i+x-1,5]
    dataset[i+x,6]=dataset[i+x,5]+dataset[i+x,4]
  }
  return(dataset)
}

Calcul_Ptcorrigé_Ecart_EcartAbsolu=function(dataset,x) {
  for (i in 1:(length(unlist(dataset[1],use.names = FALSE))-x)) {
    dataset[i+x,7]=dataset[i+x,6]*dataset[i,3]
    dataset[i+x,8]=dataset[i+x,2]-dataset[i+x,7]
    dataset[i+x,9]=abs(dataset[i+x,8])
  }
  return(dataset)
}

#library("writexl")
#write_xlsx(x,"C:\\Users\\msi\\Desktop\\test2.xlsx")


Lissage_Triple=function(Debut,alpha,beta,gamme) {
  library("readxl")
  library("writexl")
  x=read_excel(choose.files())
  x=ADDITIONDATAFRAMES(x)
  y=Calcul_Ptcorrigé_Ecart_EcartAbsolu(Calcul_St_Lt_Tt_Pt(Calcul_St_Lt_Tt_Init(x,Debut),Debut,alpha,beta,gamme),Debut)
  write_xlsx(y,"C:\\Users\\msi\\Desktop\\test6.xlsx")
  return(y)
}

Visualisation_Graphique_Lissage_Triple=function(x) {
  A=unlist(x[2],use.names = FALSE)
  B=unlist(x[7],use.names = FALSE)
  png(file = "line_chart_2_lines.jpg")
  plot(A,type = "o",col = "red", xlab = "Période", ylab = "Quantité", main = "Prévision sur la demande, Méthode Lissage triple")
  lines(B, type = "o", col = "blue")
  dev.interactive()
  dev.off()
}

require(tcltk)
Fenetre=function() {
  
  DebutVariable=tclVar("")
  AlphaVariable=tclVar("")
  BetaVariable=tclVar("")
  GammaVariable=tclVar("")
  
  tt=tktoplevel()
  tkwm.title(tt,"Prévision de la demande - Différents Méthodes")
  
  Debut.entry=tkentry(tt,textvariable=DebutVariable)
  Alpha.entry=tkentry(tt,textvariable=AlphaVariable)
  Beta.entry=tkentry(tt,textvariable=BetaVariable)
  Gamma.entry=tkentry(tt,textvariable=GammaVariable)
  
  Lissage_Triplee=function() {
    p=Lissage_Triple(as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable)),as.double(tclvalue(BetaVariable)),as.double(tclvalue(GammaVariable)))
    tkmessageBox(message=paste("Merci de trouver le resultat dans le nouveau excel cr�e dans le bureau sous le nom de test6.xlsx, veuillez aussi cr�er une copie car le document va se faire �craser dans le cas d'un nouveau usage du script"))
    return(p)
  }
  Lissage.but=tkbutton(tt, text="Méthode du Lissage Triple", command=Lissage_Triplee)
  Visu_Graphique_Triple=function() {
    Visualisation_Graphique_Lissage_Triple(Lissage_Triple(as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable)),as.double(tclvalue(BetaVariable)),as.double(tclvalue(GammaVariable))))
    tkmessageBox(message=paste("Merci de trouver le graphe dans le dossier 'Documents, veuillez �galement cr�er une copie car l'image va se faire �craser dans le cas d'un nouveau usage du script"))
  }
  VisuLissTripl.but=tkbutton(tt, text="Visualiser Les graphes Demandes/Prévisions", command=Visu_Graphique_Triple)
  
  
  reset=function(){
    tclvalue(DebutVariable)=4
    tclvalue(AlphaVariable)=0.3
    tclvalue(BetaVariable)=0.2
    tclvalue(GammaVariable)=0.1
  }
  reset.but=tkbutton(tt, text="Remettre les valeurs par défaut", command=reset)
  
  
  tkgrid(tklabel(tt,text="Veuillez choisir la période du début"), Debut.entry, pady= 10, padx= 10)
  tkgrid(tklabel(tt,text="Veuillez choisir la valeur d'Alpha"), Alpha.entry, pady= 10, padx= 10)
  tkgrid(tklabel(tt,text="Veuillez choisir la valeur de Beta"), Beta.entry, pady= 10, padx= 10)
  tkgrid(tklabel(tt,text="Veuillez choisir la valeur de Gamma"), Gamma.entry, pady= 10, padx= 10)
  tkgrid(reset.but, Lissage.but,VisuLissTripl.but, pady=10, padx=10)
  tkgrid()
  
}