convert.gis.output.to.r.format<-function(barrier=read.csv("barrier.csv")){
	
	
####  TO GET ARCGIS OUTPUT INTO THE FORMAT NEEDED TO RUN THE SCRIPT IN R  ####
# output: passability, barrier id, segment 1 (start), segment 2 (end), nat_barrier (whether or not the barrier is natural), section 1.2 (binding segment 1 and segment 2 - for use in another function)

#barrier.csv contains barrier ID, the passability for each barrier and the segments that border the barrier

# if the barrier.csv file actually has no barriers in it (i.e. barrier$Pass has 1's down the entire column),
# then we want to stop the algorithm and let the user know that their riverscape is one without natural
# or artificial barriers. 
if (length(barrier$Pass)== sum(barrier$Pass)) 
	{
	stop("
	*** There are no artificial or natural barriers in the riverscape. ***
	*** All passability values in your barrier.csv file are 1 ***
	*** Analysis will not proceed past this point. ***")
	}

else 

unique.barriers<-with(barrier, unique(Bar_ID))
#extract all of the unique barrier letters in order to later match them up with Bar_ID and get the 2 sections/segments that neighbour the barrier

segments.and.barriers<-NULL

for (i in 1:length(unique.barriers))
	{
	#find in which *position* there is a match between barrier$Bar_ID and unique barriers		
	index<-match(barrier$Bar_ID,unique.barriers[i])
	
	section.pair<-barrier$Seg_ID[!is.na(index)]
	#it tells you from the positions it found above what the barrier$Seg_ID value is
	
	#obtain the passability value by matching the barrier ID 	
	passability<-barrier$Pass[match(unique.barriers[i],barrier$Bar_ID)]
	
	#determine whether it is a natural or artificial barrier
	barrier.type<-barrier$nat_barrier[match(unique.barriers[i],barrier$Bar_ID)]
	
	#obtain the barrier id (note: you can't concatenate a factor and numbers (i.e. section name (a letter) with barrier ID and passability (numbers)), so you have to turn it into a list)
	barrier.id<-list(unique.barriers[i])
			
	sections.and.barriers<-data.frame(barrier.id,section.pair[1],section.pair[2],passability,barrier.type)
			
	names(sections.and.barriers)<-c("Bar_ID","Seg_1","Seg_2","Pass","nat_barrier")

	#use these column headings as these are the headings that Christina uses in her ArcGIS output files
	segments.and.barriers<-rbind(segments.and.barriers, sections.and.barriers)
	

	}
		
#we want to get information for both directions between segments e.g. 1 to 2 as well as the 2 to 1
#create a column where we paste the start and end segments (separated by a comma)
rev.segments.and.barriers<-segments.and.barriers
names(rev.segments.and.barriers)<-c("Bar_ID","Seg_2","Seg_1","Pass","nat_barrier")
rev.segments.and.barriers<-rev.segments.and.barriers[,c(1,3,2,4,5)]
segments.and.barriers<-rbind(segments.and.barriers,rev.segments.and.barriers)
segments.and.barriers$section1.2<-with(segments.and.barriers,paste(Seg_1,Seg_2, sep=","))

write.table(segments.and.barriers, "segments.and.barriers.csv", row.names=F, sep=",")


}