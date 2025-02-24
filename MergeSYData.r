#install.packages('RODBC')
#To comment several lines in R,use  "control + shift + C"

#Importing Followup data for Humanist
library(RODBC)
library(DBI)
con <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Temp/SY_Data/Humanist/SearchYouth.mdb")
#con <- odbcConnectAccess('C:/Temp/SY_Data/Humanist/SearchYouth.mdb')
#con <- odbcdriverConnectAccess('C:/Temp/SY_Data/Humanist/SearchYouth.mdb')
alcohol63<-sqlFetch(con, 'alcohol')
alcoholenrollment63<-sqlFetch(con, 'alcoholenrollment')
enrollment63<-sqlFetch(con, 'enrollment')
formchanges63<-sqlFetch(con, 'formchanges')
lifestage63<-sqlFetch(con, 'lifestage')
notes63<-sqlFetch(con, 'notes')
reschedule63<-sqlFetch(con, 'reschedule')
#studyexit63<-sqlFetch(con, 'studyexit')
tracking63<-sqlFetch(con, 'tracking')
vl63<-sqlFetch(con, 'vl')
withdrawal63<-sqlFetch(con, 'withdrawal')
infantoutcomes63<-sqlFetch(con, 'infantoutcomes')
maternal63<-sqlFetch(con, 'maternal')
phqsurvey63<-sqlFetch(con, 'phqsurvey')
studyexit63<-sqlFetch(con, 'studyexit')

odbcClose(con)

#Importing Followup data for Midida
library(RODBC)
con <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Temp/SY_Data/Midida/SearchYouth.mdb")
alcohol73<-sqlFetch(con, 'alcohol')
alcoholenrollment73<-sqlFetch(con, 'alcoholenrollment')
enrollment73<-sqlFetch(con, 'enrollment')
formchanges73<-sqlFetch(con, 'formchanges')
lifestage73<-sqlFetch(con, 'lifestage')
notes73<-sqlFetch(con, 'notes')
reschedule73<-sqlFetch(con, 'reschedule')
#studyexit73<-sqlFetch(con, 'studyexit')
tracking73<-sqlFetch(con, 'tracking')
vl73<-sqlFetch(con, 'vl')
withdrawal73<-sqlFetch(con, 'withdrawal')
infantoutcomes73<-sqlFetch(con, 'infantoutcomes')
maternal73<-sqlFetch(con, 'maternal')
phqsurvey73<-sqlFetch(con, 'phqsurvey')
studyexit73<-sqlFetch(con, 'studyexit')

odbcClose(con)

#Importing Followup data for Ngeri
library(RODBC)
con <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Temp/SY_Data/Ngeri/SearchYouth.mdb")
#con <- odbcConnectAccess('C:/Temp/SY_Data/Ngeri/SearchYouth.mdb')
alcohol53<-sqlFetch(con, 'alcohol')
alcoholenrollment53<-sqlFetch(con, 'alcoholenrollment')
enrollment53<-sqlFetch(con, 'enrollment')
formchanges53<-sqlFetch(con, 'formchanges')
lifestage53<-sqlFetch(con, 'lifestage')
notes53<-sqlFetch(con, 'notes')
reschedule53<-sqlFetch(con, 'reschedule')
#studyexit53<-sqlFetch(con, 'studyexit')
tracking53<-sqlFetch(con, 'tracking')
vl53<-sqlFetch(con, 'vl')
withdrawal53<-sqlFetch(con, 'withdrawal')
infantoutcomes53<-sqlFetch(con, 'infantoutcomes')
maternal53<-sqlFetch(con, 'maternal')
phqsurvey53<-sqlFetch(con, 'phqsurvey')
studyexit53<-sqlFetch(con, 'studyexit')

odbcClose(con)
#Importing Followup data for Ogongo
library(RODBC)
con <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Temp/SY_Data/Ogongo/SearchYouth.mdb")
#con <- odbcConnectAccess('C:/Temp/SY_Data/Ogongo/SearchYouth.mdb')
alcohol55<-sqlFetch(con, 'alcohol')
alcoholenrollment55<-sqlFetch(con, 'alcoholenrollment')
enrollment55<-sqlFetch(con, 'enrollment')
formchanges55<-sqlFetch(con, 'formchanges')
lifestage55<-sqlFetch(con, 'lifestage')
notes55<-sqlFetch(con, 'notes')
reschedule55<-sqlFetch(con, 'reschedule')
#studyexit55<-sqlFetch(con, 'studyexit')
tracking55<-sqlFetch(con, 'tracking')
vl55<-sqlFetch(con, 'vl')
withdrawal55<-sqlFetch(con, 'withdrawal')
infantoutcomes55<-sqlFetch(con, 'infantoutcomes')
maternal55<-sqlFetch(con, 'maternal')
phqsurvey55<-sqlFetch(con, 'phqsurvey')
studyexit55<-sqlFetch(con, 'studyexit')

odbcClose(con)

#Importing Followup data for Oyani
library(RODBC)
con <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Temp/SY_Data/Oyani/SearchYouth.mdb")
#con <- odbcConnectAccess('C:/Temp/SY_Data/Oyani/SearchYouth.mdb')
alcohol90<-sqlFetch(con, 'alcohol')
alcoholenrollment90<-sqlFetch(con, 'alcoholenrollment')
enrollment90<-sqlFetch(con, 'enrollment')
formchanges90<-sqlFetch(con, 'formchanges')
lifestage90<-sqlFetch(con, 'lifestage')
notes90<-sqlFetch(con, 'notes')
reschedule90<-sqlFetch(con, 'reschedule')
#studyexit90<-sqlFetch(con, 'studyexit')
tracking90<-sqlFetch(con, 'tracking')
vl90<-sqlFetch(con, 'vl')
withdrawal90<-sqlFetch(con, 'withdrawal')
infantoutcomes90<-sqlFetch(con, 'infantoutcomes')
maternal90<-sqlFetch(con, 'maternal')
phqsurvey90<-sqlFetch(con, 'phqsurvey')
studyexit90<-sqlFetch(con, 'studyexit')

odbcClose(con)

#Importing Followup data for Sena
library(RODBC)
con <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Temp/SY_Data/Sena/SearchYouth.mdb")
#con <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Temp/SY_Data/Sena/SearchYouth.mdb")
#con <- odbcConnectAccess('C:/Temp/SY_Data/Sena/SearchYouth.mdb')
alcohol64<-sqlFetch(con, 'alcohol')
alcoholenrollment64<-sqlFetch(con, 'alcoholenrollment')
enrollment64<-sqlFetch(con, 'enrollment')
formchanges64<-sqlFetch(con, 'formchanges')
lifestage64<-sqlFetch(con, 'lifestage')
notes64<-sqlFetch(con, 'notes')
reschedule64<-sqlFetch(con, 'reschedule')
#studyexit64<-sqlFetch(con, 'studyexit')
tracking64<-sqlFetch(con, 'tracking')
vl64<-sqlFetch(con, 'vl')
withdrawal64<-sqlFetch(con, 'withdrawal')
infantoutcomes64<-sqlFetch(con, 'infantoutcomes')
maternal64<-sqlFetch(con, 'maternal')
phqsurvey64<-sqlFetch(con, 'phqsurvey')
studyexit64<-sqlFetch(con, 'studyexit')

odbcClose(con)

#Importing Followup data for Sibuoche
library(RODBC)
con <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Temp/SY_Data/Sibuoche/SearchYouth.mdb")
#con <- odbcConnectAccess('C:/Temp/SY_Data/Sibuoche/SearchYouth.mdb')
alcohol69<-sqlFetch(con, 'alcohol')
alcoholenrollment69<-sqlFetch(con, 'alcoholenrollment')
enrollment69<-sqlFetch(con, 'enrollment')
formchanges69<-sqlFetch(con, 'formchanges')
lifestage69<-sqlFetch(con, 'lifestage')
notes69<-sqlFetch(con, 'notes')
reschedule69<-sqlFetch(con, 'reschedule')
#studyexit69<-sqlFetch(con, 'studyexit')
tracking69<-sqlFetch(con, 'tracking')
vl69<-sqlFetch(con, 'vl')
withdrawal69<-sqlFetch(con, 'withdrawal')
infantoutcomes69<-sqlFetch(con, 'infantoutcomes')
maternal69<-sqlFetch(con, 'maternal')
phqsurvey69<-sqlFetch(con, 'phqsurvey')
studyexit69<-sqlFetch(con, 'studyexit')       

odbcClose(con)

#Importing Followup data for Kisegi+Magunga
library(RODBC)
con <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Temp/SY_Data/Kisegi/SearchYouth.mdb")
#con <- odbcConnectAccess('C:/Temp/SY_Data/Kisegi/SearchYouth.mdb')
#alcohol59<-sqlFetch(con, 'alcohol')
#alcoholenrollment59<-sqlFetch(con, 'alcoholenrollment')
enrollment59<-sqlFetch(con, 'enrollment')
formchanges59<-sqlFetch(con, 'formchanges')
lifestage59<-sqlFetch(con, 'lifestage')
notes59<-sqlFetch(con, 'notes')
reschedule59<-sqlFetch(con, 'reschedule')
#studyexit59<-sqlFetch(con, 'studyexit')
tracking59<-sqlFetch(con, 'tracking')
vl59<-sqlFetch(con, 'vl')
withdrawal59<-sqlFetch(con, 'withdrawal')
infantoutcomes59<-sqlFetch(con, 'infantoutcomes')
maternal59<-sqlFetch(con, 'maternal')
phqsurvey59<-sqlFetch(con, 'phqsurvey')
studyexit59<-sqlFetch(con, 'studyexit')       

odbcClose(con)

#Importing Followup data for Nyatoto+Nyadenda
library(RODBC)
con <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Temp/SY_Data/Nyatoto/SearchYouth.mdb")
#con <- odbcConnectAccess('C:/Temp/SY_Data/Nyatoto/SearchYouth.mdb')
#alcohol50<-sqlFetch(con, 'alcohol')
#alcoholenrollment50<-sqlFetch(con, 'alcoholenrollment')
enrollment50<-sqlFetch(con, 'enrollment')
formchanges50<-sqlFetch(con, 'formchanges')
lifestage50<-sqlFetch(con, 'lifestage')
notes50<-sqlFetch(con, 'notes')
reschedule50<-sqlFetch(con, 'reschedule')
#studyexit50<-sqlFetch(con, 'studyexit')
tracking50<-sqlFetch(con, 'tracking')
vl50<-sqlFetch(con, 'vl')
withdrawal50<-sqlFetch(con, 'withdrawal')
infantoutcomes50<-sqlFetch(con, 'infantoutcomes')
maternal50<-sqlFetch(con, 'maternal')
phqsurvey50<-sqlFetch(con, 'phqsurvey')
studyexit50<-sqlFetch(con, 'studyexit')       

odbcClose(con)

#Importing Followup data for Ongo+Othoro+Uriri
library(RODBC)
con <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Temp/SY_Data/Ongo/SearchYouth.mdb")
#con <- odbcConnectAccess('C:/Temp/SY_Data/Ongo/SearchYouth.mdb')
#alcohol67<-sqlFetch(con, 'alcohol')
#alcoholenrollment67<-sqlFetch(con, 'alcoholenrollment')
enrollment67<-sqlFetch(con, 'enrollment')
formchanges67<-sqlFetch(con, 'formchanges')
lifestage67<-sqlFetch(con, 'lifestage')
notes67<-sqlFetch(con, 'notes')
reschedule67<-sqlFetch(con, 'reschedule')
#studyexit67<-sqlFetch(con, 'studyexit')
tracking67<-sqlFetch(con, 'tracking')
vl67<-sqlFetch(con, 'vl')
withdrawal67<-sqlFetch(con, 'withdrawal')
infantoutcomes67<-sqlFetch(con, 'infantoutcomes')
maternal67<-sqlFetch(con, 'maternal')
phqsurvey67<-sqlFetch(con, 'phqsurvey')
studyexit67<-sqlFetch(con, 'studyexit')       

odbcClose(con)

#Importing Followup data for Kitare
library(RODBC)
con <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Temp/SY_Data/Kitare/SearchYouth.mdb")
#con <- odbcConnectAccess('C:/Temp/SY_Data/Kitare/SearchYouth.mdb')
#alcohol56<-sqlFetch(con, 'alcohol')
#alcoholenrollment56<-sqlFetch(con, 'alcoholenrollment')
enrollment56<-sqlFetch(con, 'enrollment')
formchanges56<-sqlFetch(con, 'formchanges')
lifestage56<-sqlFetch(con, 'lifestage')
notes56<-sqlFetch(con, 'notes')
reschedule56<-sqlFetch(con, 'reschedule')
#studyexit56<-sqlFetch(con, 'studyexit')
tracking56<-sqlFetch(con, 'tracking')
vl56<-sqlFetch(con, 'vl')
withdrawal56<-sqlFetch(con, 'withdrawal')
infantoutcomes56<-sqlFetch(con, 'infantoutcomes')
maternal56<-sqlFetch(con, 'maternal')
phqsurvey56<-sqlFetch(con, 'phqsurvey')
studyexit56<-sqlFetch(con, 'studyexit')       

odbcClose(con)

library("plyr")
#Merging the data from various facilities
totalal <- rbind.fill(alcohol63, alcohol73, alcohol53, alcohol55, alcohol90, alcohol64, alcohol69)

totalalcen <- rbind.fill(alcoholenrollment63, alcoholenrollment73,  alcoholenrollment53,  
                         alcoholenrollment55, alcoholenrollment90, alcoholenrollment64, 
                         alcoholenrollment69)

totalen <- rbind.fill(enrollment63, enrollment73, enrollment53, enrollment55, enrollment90, 
                      enrollment64, enrollment69, enrollment56)

totalfc <- rbind.fill(formchanges63, formchanges73,  formchanges53,  formchanges55, formchanges90, 
                      formchanges64, formchanges69, formchanges59, formchanges56)

totalls <- rbind.fill(lifestage63, lifestage73, lifestage53, lifestage55, lifestage90, 
                      lifestage64, lifestage69, lifestage59, lifestage56)

totaln <- rbind.fill(notes63, notes73, notes53, notes55, notes90, notes64, notes69, 
                     notes59, notes56)

totalre <- rbind.fill(reschedule63, reschedule73, reschedule53, reschedule55, reschedule90, 
                      reschedule64, reschedule69, reschedule59, reschedule56)

#totalex <- rbind.fill(studyexit63, studyexit73, studyexit53, studyexit55, studyexit90, studyexit64, studyexit69, studyexit59)

totaltr <- rbind.fill(tracking63, tracking73, tracking53, tracking55, tracking90, tracking64, 
                      tracking69, tracking59, tracking56)

totalvl <- rbind.fill(vl63, vl73, vl53, vl55, vl90, vl64, vl69, vl59, vl56)

totalwd <- rbind.fill(withdrawal63, withdrawal73, withdrawal53, withdrawal55, withdrawal90, 
                      withdrawal64, withdrawal64, withdrawal69, withdrawal59, withdrawal56)

totalinf<-rbind.fill(infantoutcomes63, infantoutcomes73, infantoutcomes53, infantoutcomes55, 
                     infantoutcomes90, infantoutcomes64, infantoutcomes69, infantoutcomes59, 
                     infantoutcomes67, infantoutcomes50, infantoutcomes56)

totalmat<-rbind.fill(maternal63, maternal73, maternal53, maternal55, maternal90, maternal64, 
                     maternal69, maternal59, maternal67, maternal50, maternal56)
 
totalphq<-rbind.fill(phqsurvey63, phqsurvey73, phqsurvey53, phqsurvey55, phqsurvey90, 
                     phqsurvey64, phqsurvey69, phqsurvey59, phqsurvey67, phqsurvey50, phqsurvey56) 

totalext<-rbind.fill(studyexit63, studyexit73, studyexit53, studyexit55, studyexit90, 
                     studyexit64, studyexit69, studyexit59, studyexit67, studyexit50, 
                     studyexit56)

#To export merged SY_DB tablets to excel - csv format
library("writexl")
write.csv( totalal,"C:/Temp/SY_Data/Merged_Data//alcohol.csv")
write.csv( totalalcen,"C:/Temp/SY_Data/Merged_Data//alcoholenrollment.csv")
write.csv( totalen,"C:/Temp/SY_Data/Merged_Data//enrollment.csv")
write.csv( totalfc,"C:/Temp/SY_Data/Merged_Data/formchanges.csv")
write.csv( totalls,"C:/Temp/SY_Data/Merged_Data/lifestage.csv")
write.csv( totaln,"C:/Temp/SY_Data/Merged_Data/notes.csv")
write.csv( totalre,"C:/Temp/SY_Data/Merged_Data/reschedule.csv")
#write.csv( totalex,"C:/Temp/SY_Data/Merged_Data/studyexit.csv")
write.csv( totaltr,"C:/Temp/SY_Data/Merged_Data/tracking.csv")
write.csv( totalvl,"C:/Temp/SY_Data/Merged_Data/vl.csv")
write.csv( totalwd,"C:/Temp/SY_Data/Merged_Data/withdrawal.csv")
write.csv(totalinf, "C:/Temp/SY_Data/Merged_Data/infantoutcomes.csv")
write.csv(totalmat, "C:/Temp/SY_Data/Merged_Data/maternal.csv")
write.csv(totalphq, "C:/Temp/SY_Data/Merged_Data/phqsurvey.csv")
write.csv(totalext, "C:/Temp/SY_Data/Merged_Data/studyexit.csv")
