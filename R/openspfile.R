

openspfile<-function(spfile){
  file.df<-readWorksheetFromFile(spfile, sheet=1,endCol = 35, forceConversion=TRUE)
  file.df<-as.data.frame(file.df)
  sp.df<-file.df[8:nrow(file.df)-1,]
  
  sp.ncol<-ncol(file.df)
  
  if(sp.ncol>=31){
    
    print(c(file.df[5,11], file.df[5,16], file.df[5,19],file.df[5,26],file.df[5,29]))
    
    sp.df<-sp.df[,c(1,7,11,14,15,
                    16,17,18,
                    19,23,24,
                    26,27,28,
                    29,30,31)]
    colnames(sp.df)<-c("num.registro","evaluado", "compciu.punt", "compciu.nivel", "compciu.quin",
                       "comuesc.punt", "comuesc.nivel", "comuesc.quin", 
                       "ingl.punt", "ingl.nivel", "ingl.quin",
                       "lectucrit.punt", "lectucrit.nivel", "lectucrit.quin",
                       "razcuant.punt", "razcuant.nivel", "razcuant.quin")
    sp.df$periodo<-file.df[1,9]
    sp.df$municipio<-file.df[3,13]
    sp.df$programa<-file.df[4,13]
    sp.df$grupo.ref<-file.df[4,5]
    return(as.data.frame(sp.df))
  }
  
  
  if(sp.ncol==28){
    
    print(c(file.df[5,11], file.df[5,16], file.df[5,19], file.df[5,26]))
    
    sp.df2<-sp.df[,c(1,7,11,14,15,
                     16,17,18,
                     19,23,24,
                     26,27,28)]
    colnames(sp.df2)<-c("num.registro","evaluado", "compciu.punt", "compciu.nivel", "compciu.quin",
                        "ingl.punt", "ingl.nivel", "ingl.quin",
                        "razcuant.punt", "razcuant.nivel", "razcuant.quin")
    sp.df<-data.frame(matrix(ncol = 17, nrow = nrow(sp.df2)))
    colnames(sp.df)<-c("num.registro","evaluado", "compciu.punt", "compciu.nivel", "compciu.quin",
                       "comuesc.punt", "comuesc.nivel", "comuesc.quin", 
                       "ingl.punt", "ingl.nivel", "ingl.quin",
                       "lectucrit.punt", "lectucrit.nivel", "lectucrit.quin",
                       "razcuant.punt", "razcuant.nivel", "razcuant.quin")
    sp.df[,1:5]<-sp.df2[,1:5]
    sp.df[,9:11]<-sp.df2[,6:8]
    sp.df[,15:17]<-sp.df2[,9:11]
    sp.df$periodo<-file.df[1,9]
    sp.df$municipio<-file.df[3,13]
    sp.df$programa<-file.df[4,13]
    sp.df$grupo.ref<-file.df[4,5]
    return(as.data.frame(sp.df))
  }
  
}