#-------------------------------------------------------------------------------
#  Developed by: Kevin A. Rodberg, Science Supervisor 
#                 Resource Evaluation Section, Water Supply Bureau, SFWMD
#                 (561) 682-6702
#  November 2019
#
#  Script is provided to import wetland data entry forms data for DMIT
#  and extract specific fields to compile data set from collection 
#  of spreadsheets
#
#-------------------------------------------------------------------------------

#--
#   package management:
#     provide automated means for first time use of script to automatically 
#	  install any new packages required for this code, with library calls 
#	  wrapped in a for loop.
#--
list.of.pkgs <-  c("readr","dplyr","reshape2", "data.table",
                   "readxl","janitor","tictoc","tcltk")

new.pkgs <- list.of.pkgs[!(list.of.pkgs %in% installed.packages()[, "Package"])]

if (length(new.pkgs)){ install.packages(new.pkgs) }
for (pkg in list.of.pkgs){ library(pkg,character.only = TRUE) }

#------------------------------------------------------
#	Define working directories and input files to read
#------------------------------------------------------
baseDir = "//ad.sfwmd.gov/dfsroot/data/wsd/SUP/devel/source/R/DMITworksheets/"
workdir ="/Input"
Outdir = "/Output"

workdir<-tclvalue(tkchooseDirectory(initialdir=baseDir,
                                    title='Select location for input spreadsheets'))
setwd(workdir)
workOutdir<-tclvalue(tkchooseDirectory(initialdir=baseDir,
                                       title='Select location for output file'))

options(digits=9)
xlFiles <-list.files(pattern = "*.xlsm")
xlFiles <-xlFiles[!grepl(paste0("~", collapse = "|"), xlFiles)]
Wetland<- as.data.frame(xlFiles)
names(Wetland)<- c('Filename')
Wetland$Filename<- as.character(Wetland$Filename)
Wetland$DMIT_ID <-NA
Wetland$DMIT_Name <-NA

WetlandList<- vector(mode = "list", length = length(xlFiles))
fileNum = 0
for (file in xlFiles){
  fileNum = fileNum + 1
  sheets <- excel_sheets(file)
  cat(paste0('\n',file,'::\n'))
  cat(paste(list(sheets[1:7]),'\n\n'))
  noMesg <- suppressMessages
  tic(paste("Read 7 tabs for",file,'\n'))
  df1 <- noMesg(read_excel(file, sheet=sheets[1], col_names = FALSE))
  df2 <- noMesg(read_excel(file, sheet=sheets[2], n_max=200, col_names=FALSE))
  df3 <- noMesg(read_excel(file, sheet=sheets[3], n_max=200, col_names=FALSE))
  df4 <- noMesg(read_excel(file, sheet=sheets[4], n_max=200, col_names=FALSE))
  df5 <- noMesg(read_excel(file, sheet=sheets[5], n_max=200, col_names=FALSE))
  df6 <- noMesg(read_excel(file, sheet=sheets[6], n_max=200, col_names=FALSE))
  df7 <- noMesg(read_excel(file, sheet=sheets[7], n_max=200, col_names=FALSE))
  readTime<-toc()
  if ((readTime$toc-readTime$tic)> 7){
    cat(paste('->>>Excel file processing was slow for',
              file,'\n->>>It is mostlikely open in Excel\n\n'))
  }
  
#------------------------------------------------------
#  Process Hydrological Tab  Sheets[1]
#------------------------------------------------------
  Wetland[Wetland$Filename==file,]$DMIT_ID<-as.numeric(df1[1,2])
  Wetland[Wetland$Filename==file,]$DMIT_Name<-as.character(df1[1,7])

  TranRowOffset = 25  # Transect blocks take up 25 rows starting on row 2
  PointOffset = 5     # Point records take up 5 rows/Point: 4 pnts/transect
  
  TransectList<-vector(mode="list",length=3)   # empty list of 3 transects
  TransectPoints<-vector(mode="list",length=4) # empty list of 4 points/transect
  for (Transect in c(1,2,3)){
    trans = NULL
    trans$TransectType<-'Hydrologic'
    trans$TansectNum<-Transect
    trans$Date<-janitor::excel_numeric_to_date(as.numeric(df1[3,2]))
    trans$FieldTeam<-as.character(df1[4,2])
    TransectList[[Transect]] <- as.data.frame(trans)
    Points <- vector(mode = "list", length = 4)
    for (Indicator in c(1,2,3,4)){
      point=NULL
      irow= 7+((Transect-1)*TranRowOffset)+((Indicator-1)*PointOffset)
      point$PointType<- as.character(df1[irow,3])
      point$TansectNum<-Transect
      point$IndicatorNum<-Indicator
      irow = 10+((Transect-1)*TranRowOffset)+((Indicator-1)*PointOffset)
      point$Lat_DecDeg<-as.numeric(df1[irow,5])
      point$Long_DecDeg<-as.numeric(df1[irow+1,5])
      #---
      #  Validate or correct data for commmon coordinate problem
      #---
      if (!is.na(point$Long_DecDeg) & point$Long_DecDeg > 0){
        cat(paste('Read Longitude =',point$Long_DecDeg,
                  '---- Western hemisphere longitude values must be negative.',
                  ' Value Corrected\n'))
        point$Long_DecDeg = -point$Long_DecDeg
      }
      point$Elevation_ftNAVD88<-as.numeric(df1[irow,11])
      Points[[Indicator]]<-data.frame(point)
    }
    TransectPoints[[Transect]]<-do.call("rbind",Points)
    names(TransectPoints[[Transect]])<-c("PointType","TansectNum",
                                         "IndicatorNum",
                                         "Lat_DecDeg","Long_DecDeg",
                                         "Elevation_ftNAVD88")
  }
  TranPntDF<-do.call("rbind",TransectPoints)
  transDF<-do.call("rbind",TransectList)
  HydrologicDF<-  merge(transDF,TranPntDF)
  #------------------------------------------------------
  # Process Soils Tabs  sheets[c(2,3,4)]
  #------------------------------------------------------
  TranIDrow =     c(4,37,69,102)    # Row offsets defined since PointOffset
  HydricSoilRow = c(20,52,84,117)   # would not be consistent 
  IndicatorRow =  c(22,54,86,119)   # per transect
  TransectList<- vector(mode = "list", length = 3)
  TransectPoints<- vector(mode = "list", length = 3)
  Transect = 0
  for (tab in sheets[c(2,3,4)]){
    Transect = Transect + 1
    if (tab == "Soil T1"){df <- df2}
    if (tab == "Soil T2"){df <- df3}
    if (tab == "Soil T3"){df <- df4}
    trans = NULL
    trans$TransectType<-"Soil"
    Transect4Point1<-as.numeric(df[TranIDrow[1],2])
    #---
    # Check and/or correct if Transect Number is different 
    # than the spreadsheet tab identification
    #---
    if (as.numeric(Transect) != Transect4Point1){
      cat(paste('\tExpecting Transect',as.numeric(Transect),
                'but found',Transect4Point1,'identified on ',tab,'\n') )
      trans$TansectNum<-Transect
    } else {
      trans$TansectNum<-Transect4Point1
    }
    trans$Date<-janitor::excel_numeric_to_date(as.numeric(df[3,2]))
    trans$FieldTeam<-as.character(df[3,5])
    TransectList[[trans$TansectNum]] <- as.data.frame(trans)
    Points <- vector(mode = "list", length = 4)
    soilIndex = 0
    for (Indicator in c(1,2,3,4)){
      for (presence in c(1,2,3,4,5,6,7,8)){
        soilIndex = soilIndex + 1
        point=NULL
        irow= 6+(PointOffset[Indicator])
        if(!is.na(df[TranIDrow[Indicator]+1,2])){
          PointType <- as.character(df[TranIDrow[Indicator]+1,2])
          if (PointType == 'Other') {
            PointType <- as.character(df[TranIDrow[Indicator]+2,2]) 
          }
          Transect4Point<-df[TranIDrow[Indicator],2]
          #---
          # Check and/or correct if Transect Number is different 
          # than the spreadsheet tab identification
          #---
          if(!is.na(Transect4Point)){
            if (as.numeric(Transect) != as.numeric(Transect4Point)){
              cat(paste('\tTransect',as.numeric(Transect4Point),
                        'identified on ',tab,'for point',Indicator,'\n') )
              point$TansectNum<-Transect
            } else {
              point$TansectNum<-as.numeric(Transect4Point)
            }
            point$PointType<-PointType
            point$IndicatorNum<-Indicator
            point$Lat_DecDeg<-as.numeric(df[TranIDrow[Indicator]+2,5])
            point$Long_DecDeg<-as.numeric(df[TranIDrow[Indicator]+3,5])
            #---
            #  Validate or correct data for commmon coordinate problem
            #---            
            if (!is.na(point$Long_DecDeg) & point$Long_DecDeg > 0){
              cat(paste('Read Longitude =',point$Long_DecDeg,
                  '---- Western hemisphere longitude values must be negative.',
                  ' Value Corrected\n'))
              point$Long_DecDeg = -point$Long_DecDeg
            }
            point$Elevation_ftNAVD88<-as.numeric(df[TranIDrow[Indicator]+2,11])
            HydricSoilIndicators <-as.character(df[HydricSoilRow[Indicator],2])
            presenceOffset=presence-1
            IndicatorParam<-as.character(df[IndicatorRow[Indicator]+
                                              presenceOffset,11])
            #---
            # Compile up to 8 Indicator Presence params and values for 
            # each point with Hydric Soils Indicator(s) Present  == Yes
            #---
            if (!is.na(HydricSoilIndicators) & 
                HydricSoilIndicators == 'Yes' &
                !is.na(IndicatorParam)){
              point$Parameter<- IndicatorParam
              #---
              #  Beginning Depth and End Depth are merged 
              #  into single Value text string
              #---
              Value <- paste0('Depth: ',
                              df[IndicatorRow[Indicator]+presenceOffset,13],
                              ' to ',
                              df[IndicatorRow[Indicator]+presenceOffset,14],
                              ' inches')
              point$Value<- Value
              Points[[soilIndex]]<-data.frame(point)
              pointNames <-names(data.frame(point))
            } else {
              soilIndex = soilIndex - 1
            }
          }
        }
      }
    }
    TransectPoints[[trans$TansectNum]]<-do.call("rbind",Points)
  }
  TranPntDF<-do.call("rbind",TransectPoints)
  transDF<-do.call("rbind",TransectList)
  SoilsDF<-  merge(transDF,TranPntDF)
  #------------------------------------------------------
  # Process Vegetation Tabs  sheets[c(5,6,7)]
  #------------------------------------------------------
  PointOffset = 45   # Transect blocks take up 45 rows starting on row 2
  TransectList<- vector(mode = "list", length = 3)
  TransectPoints<- vector(mode = "list", length = 4)
  Transect = 0
  for (tab in sheets[c(5,6,7)]){
    Transect = Transect + 1
    # for (tab in sheets[c(5)]){
    if (tab == "Vegetation T1"){df <- df5}
    if (tab == "Vegetation T2"){df <- df6}
    if (tab == "Vegetation T3"){df <- df7}
    trans = NULL
    trans$TransectType<-"Vegetation"
    Transect4Point1<-as.numeric(df[4,2])
    #---
    # Check and/or correct if Transect Number is different 
    # than the spreadsheet tab identification
    #---
    if (as.numeric(Transect) != Transect4Point1){
      cat(paste('\tExpecting Transect',as.numeric(Transect),
                'but found',Transect4Point1,'identified on ',tab,'\n') )
      trans$TansectNum<-Transect
    } else {
      trans$TansectNum<-Transect4Point1
    }
    trans$Date<-janitor::excel_numeric_to_date(as.numeric(df[3,2]))
    trans$FieldTeam<-as.character(df[3,5])
    TransectList[[trans$TansectNum]] <- as.data.frame(trans)
    Points <- vector(mode = "list", length = 48)
    vegIndex = 0
    for (Indicator in c(1,2,3,4)){
      irow= 3+((Indicator-1)*PointOffset)
      PointDesc <- df[irow+2,2]
      Transect4Point<-df[irow+1,2]
      if(nrow(SoilsDF[SoilsDF$TansectNum == Transect &
                      SoilsDF$IndicatorNum == Indicator ,]) ==0 &
         !is.na(PointDesc)){
          cat(paste('\tSoils Transect Point location not available to ',
                    'provide coordinates for Veg: T', Transect,'-',Indicator,'\n'))
      }
      #---
      # Check and/or correct if Transect Number is different 
      # than the spreadsheet tab identification
      #---
      if(!is.na(Transect4Point)){
        if (as.numeric(Transect) != as.numeric(Transect4Point)){
        cat(paste('\t\tTransect',as.numeric(Transect4Point),
                  'identified on ',tab,'for point',Indicator,'Redefined to',
                  Transect,'\n') )
        }
      }
      paramIndex = 0
      colNum = 7
      for (param in c('Canopy','Sub-Canopy','Groundcover')){
        paramIndex = paramIndex + 1
        statusIndex = 0
        for (status in c('OBL','FACW','FAC','UPL')){
          statusIndex= statusIndex + 1
          vegIndex= vegIndex + 1
          colNum = colNum + 1
          point=NULL
          #---
          # Compile Total Stratum % by Status as params and values for   
          # each point if Transect number present for a point 
          #---
          if(!is.na(Transect4Point)){
            point$PointType<- as.character(PointDesc)
            if (as.numeric(Transect) != as.numeric(Transect4Point)){
              point$TansectNum<-Transect
            } else {
              point$TansectNum<-as.numeric(Transect4Point)
            }
            point$IndicatorNum<-Indicator
            irow= 41+((Indicator-1)*PointOffset)
            if(nrow(SoilsDF[SoilsDF$TansectNum == Transect &
                            SoilsDF$IndicatorNum == Indicator,]) >0){
              #---
              # Identify Soils point record matching Vegetation pointrecord
              #---
              soilsRec <- SoilsDF[SoilsDF$TansectNum == Transect & 
                                    SoilsDF$IndicatorNum == Indicator,][1,]
              point$Lat_DecDeg<-soilsRec$Lat_DecDeg
              point$Long_DecDeg<-soilsRec$Long_DecDeg
              point$Elevation_ftNAVD88<-soilsRec$Elevation_ftNAVD88
            } else {
              point$Lat_DecDeg<-NA
              point$Long_DecDeg<-NA
              point$Elevation_ftNAVD88<-NA
            }
            point$Parameter<- paste0(param,'_',status)
            Value <-as.character(df[irow,colNum])
            point$Value<- Value
            Points[[vegIndex]]<-data.frame(point)
            pointNames <-names(data.frame(point))
          }
        }
      }
    }
    TransectPoints[[trans$TansectNum]]<-do.call("rbind",Points)
    names(TransectPoints[[trans$TansectNum]])<-pointNames
  }
  TranPntDF<-do.call("rbind",TransectPoints)
  transDF<-do.call("rbind",TransectList)
  VegetationDF<-  merge(transDF,TranPntDF)
  #------------------------------------------------------
  # Merge Data from all Tabs
  #------------------------------------------------------
  AllTabs<-Reduce(function(...) merge(..., all = TRUE),
                  list(HydrologicDF, SoilsDF, VegetationDF))
  AllTabs<-AllTabs[(substr(AllTabs$TransectType,1,3)=='Veg' & 
                      !is.na(AllTabs$PointType))
                   | (substr(AllTabs$TransectType,1,3)=='Hyd' & 
                        !is.na(AllTabs$Lat_DecDeg))
                   | (substr(AllTabs$TransectType,1,3)=='Soi'),]
  WetlandList[[fileNum]]<-merge(Wetland[fileNum,],AllTabs)
}
WetlandsDF<-do.call("rbind",WetlandList)
write.csv(WetlandsDF,paste0(workOutdir,'/DMITsites.csv'))
cat(paste('Output has been saved to:',
          paste0(workOutdir,'/DMITsites.csv')),'\n')
