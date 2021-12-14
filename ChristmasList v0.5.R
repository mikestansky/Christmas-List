##################################################################################################
#Author:Michael Stansky
#Christmas Exchange - 2017-12-01
#This program reads in an excel file with 3 tabs: XmasList, ExcludePair and ForcedPair
#It creates a directed graph of to/from pairs for a gift exchange.
#ExcludePair and ForcedPair are lists of constraints to exclude from the list (i.e. Mike cannot give
# a gift to Amy) and to Force (e.i. Mike must be paired with Tony since he already bought the gift).
#ExcludePair and ForcedPair is one direction, you must create two records for bi-directional exclusions.
#Since the exchange list creates a non-deterministic directed graph, there may be several closed rings of 
#giving.  You can set limits on the number of rings and minimum ring size.  However, if the excluded and forced
#pairings make an impossible list, maxiter sets a stopping criteria.
#
#The output is a file with the to/from list and an image of the directed graph.
##################################################################################################

#install.packages("magick")

#Enter file path with Excel file list of names
filepath="C:\\Users\\mstansky\\OneDrive\\Documents\\Xmas\\"

#x<-winDialog(type = c("okcancel"), "Save Excel File First")

require(dplyr) 
require(xlsx)
require(igraph)
require(magick)

#Read in the list and the excluded and forced pairings
nameslist<-read.xlsx(paste(trimws(filepath),"Xmas List.xlsx",sep=""),sheetName="XmasList")
excludepair<-read.xlsx(paste(trimws(filepath),"Xmas List.xlsx",sep=""),sheetName="ExcludePair")
forcedpair<-read.xlsx(paste(trimws(filepath),"Xmas List.xlsx",sep=""),sheetName="ForcedPair")




#Set limiting parameters
max_num_rings=1   #number of allowable rings
min_ring_size=3   #if more than one ring is allowed, how big should the smallest ring be
maxiter=100       #in case no solution can be found, how long should the search iterate


##################################################################################################
##Instead of the excel file, you can uncomment this section, and use it for the names list
#nameslist<-data.frame(name=c("Mike","Amy","Eric","Tony","Michelle","Carie","Greg", "Marren",
#           "Carol","Mark","Beth","Anne","Richard","Jackie","Josh","Chuck"))
#
#
#excludepair<-data.frame(rbind(c("Mike","Amy"),c("Amy","Mike"),c("Richard","Jackie"),c("Marren","Tony")))
#names(excludepair)<-c("From","To")
#
#forcedpair<-data.frame(rbind(c("Carie","Amy"),c("Marren","Michelle"),c("Mark","Anne"),c("Anne","Carie")))
#names(forcedpair)<-c("From","To")
#
##################################################################################################



#pull off self matches and make sure that no one matches themselves
selfmatch=cbind(nameslist,nameslist)
names(selfmatch)<-c("From","To")

# combine selfmatches with excluded list, rename columns for later use
fexcludepair<-data.frame(rbind(excludepair,selfmatch))
fexcludepair<-na.omit(fexcludepair)
names(fexcludepair)<-c("santa","recipient")

fexcludepair
#turn forced to and from into frames for easier handling
forcedfrom<-data.frame(Name=forcedpair[,1])
forcedto<-data.frame(Name=forcedpair[,2])

#find the names that are not in forced pairs
freefrom<-anti_join(nameslist,forcedfrom)
freeto<-anti_join(nameslist,forcedto)
names(forcedpair)<-c("santa","recipient")



#initialize while loop break criteria
goodrun<-FALSE; i=1;
#loop though possible matching an make sure there are no bad matches
while (goodrun==FALSE) {
  #create list of random pairs
  newlist<-cbind(sample_n(freefrom,nrow(freefrom)),sample_n(freeto,nrow(freeto)))
  names(newlist)<-c("santa","recipient")
  newlist
    
  matchednames<-data.frame(rbind(newlist,forcedpair))
  #str(matchednames)
  
  #count the number of bad matches, either self matches or on the exception list
  badmatches<-semi_join(matchednames,fexcludepair,by = c("santa", "recipient"))
  nrow(badmatches)
  
  #count the number of rings and min ring size in graph
  net <- graph_from_data_frame(d=matchednames, vertices=nameslist, directed=T) 
  netring<-components(net)$no
  smallestring<-min(components(net)$csize)
  
  
  #if there are no bad matches and the ring is less than the max, then break the loop
  if (nrow(badmatches)==0 & netring<=max_num_rings & smallestring>=min_ring_size) {goodrun<-TRUE}
  
  #if a valid network can't be found within maxiter tries, the loop is broken
  if (i==maxiter){
    print("No solution was found, try increasing max_num_rings or decreasing min_ring_size")
    goodrun<-FALSE
    break
  }
  i=i+1
}



#plot the network and more to an image for export
  img <- image_graph(600, 600, res = 96)
  p<-plot(net, vertex.shape="none", vertex.label=V(net)$santa, 
          vertex.label.font=2.5, vertex.label.color="dark green",
          vertex.label.cex=1, edge.color="red")
    
  print(p)
  dev.off()  
  img<-img %>% image_annotate(" Merry Christmas ", size = 30, color = "red", boxcolor = "green", location = "+180+20")
  img

#export the matches and graph as an image
  matchednames2<-matchednames %>% mutate(text=paste(santa,"will give a gift to",recipient))
  write.xlsx(matchednames2,paste(trimws(filepath),"Output Xmas List 2021.xlsx",sep=""),sheetName="XmasList")
  image_write(img, path = paste(trimws(filepath),"Output Xmas List 2021.jpg",sep=""),   format = "jpg")
  print(matchednames2$text)


 
#image_annotate("Merry", size = 30, color = "green", boxcolor = "red", location = "+270+240")   %>% 
#  image_annotate("Christmas", size = 30, color = "red", boxcolor = "green", location = "+240+290")
  
