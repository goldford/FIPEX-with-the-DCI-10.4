sum.fx<- function(adj.matrix,passability,lengths) {
	
#function(adj.matrix, passability) calls on "adj.matrix" and "passability" that were created in previous functions.  When you call a function, it creates a bunch of objects, but these objects disappear when you open a new function.  In order to be able to use these objects, you can either "assign(object.name)", or you can do what was done above function(object.names).  Because all of these functions are called in "dci.fxs.r" and "dci.fxs.theo.restoration.r", the information is passed on from one function to another.
	
	
# WHAT THIS FUNCTION DOES: it creates all the data for a summary table (start and end segments, pathway, barriers it goes through, passability, length of start and end segments)
# BUT this dataframe is NOT outputted into Excel in this function.  It's returned and then called in by graph.and.data.setup.for.DCI.r.


lengths<-read.csv("length.csv")


x2<-NULL #section start
y2<-NULL #section end
start.section.length<-NULL
finish.section.length<-NULL
path2<-NULL
pathway.pass<-NULL
barrier.id<-NULL			

# obtain the section names from the adj matrix
sections<-rownames(adj.matrix)


# there is a problem with only one barrier, since sp.between requires a list, and we provide a vector.
# A quick work-aroud is to manually calcualte the DCI with only on barrier and skip the more complicated
# steps below
if(length(sections)>2)
{
	
g2<-new("graphAM",adjMat=adj.matrix, edgemode="directed")



#get a list of adjacent sections (i.e. find all possible pathways that exist in the riverscape)
for (i in 1:length(sections))
{
	#need it to look through the matrix, "i" cycles down the columns and "j" cycles across the rows
	for (j in 1:length(sections))
	{

		x<-sections[i]
		y<-sections[j]

		path.all<-sp.between(g=g2,start=x,finish=y)
		#this extracts every possible pair that exists in the matrix
		#sp.between = shortest path between 2 pairs
				
		path<-path.all[[1]]$path_detail
		#this pulls out the path information for each pair of sections
		#e.g. [1] "1" "2" "3" - to go from 1 to 3 you must go through 2
		
		path.length<-length(path)
		#we need to get the length so that the k-loop (below) can pull out the appropriate barrier information 
		
		#now that we have all of the possible combinations that exist between sections, we now need to find the barriers that exist bewteen these sections
		
		x2<-c(x2,x)
		y2<-c(y2,y)
			
		start.section<-path[1]
		#this should grab the 1st segment of the path
		finish.section<-path[path.length]	
		#this should grab the last segment of the path
			
		#need to get the length of the start and finish segments of the pathway
		new.start.length<-lengths$Shape_Length[match(start.section, lengths$Seg_ID)]
		new.finish.length<-lengths$Shape_Length[match(finish.section, lengths$Seg_ID)]
			
		start.section.length<-c(start.section.length, new.start.length)
		finish.section.length<-c(finish.section.length, new.finish.length)
		
		section1.2<-NULL
		section.1<-NULL
		section.2<-NULL
		
		if (path.length<2) section1.2<-NA 
		#this will assign "NA"s to 1-1, 2-2, 3-3, etc...
		
		else
		for(k in 1:(path.length-1))
			#use "-1" because we are looking at the number of barriers between each section.  number of barriers is = to number of sections - 1.
			{
				section1<-path[k]
				section2<-path[k+1]
				#k+1 is used so that R knows to go to the 2nd element in path
				
				section.1<-c(section.1, section1)
				section.2<-c(section.2, section2)
						
				section.paste<-paste(section1,section2,sep=",")
				section1.2<-c(section1.2, section.paste)
				#need to create this new column, section1.2, in order to match it to passability$section1.2 - see function "graph.fx.r ### INPUT DATA ###"
			}	
		#find the barriers that are between each pair of sections
		all.barriers<-passability$Bar_ID[match(section1.2,passability$section1.2)]
		
		#give a list of all of the passability values from start to end of the path		
		new.barrier.pass2<-passability$Pass[match(all.barriers,passability$Bar_ID)]
					
		#need to get the product for each new.barrier.pass$all.passabilities
		#add a new column to the new.barrier.pass dataframe
		new.barrier.pass<-prod(new.barrier.pass2)
		pathway.pass<-c(pathway.pass,new.barrier.pass)
			
		#now we need to turn these objects into 1 vector in order to use cbind to make these objects into 1 dataframe
		
		new.barrier.id<-paste(all.barriers,sep="",collapse=",")
		barrier.id<-c(barrier.id, new.barrier.id)
					
		path1<-paste(path,collapse=",")
		#we want to get "a" "b" "c" to be 1 character (i.e. "a,b,c")
		#this allows it to be recognized as 1 element in the dataframe
		path2<-c(path2,path1)
		}

	sum.table<-cbind(x2, y2, path2, barrier.id, pathway.pass, start.section.length, finish.section.length)
	sum.table<-as.data.frame(sum.table)
	sum.table$pathway.pass<-as.numeric(as.character(sum.table$pathway.pass))
	sum.table$start.section.length<-as.numeric(as.character(sum.table$start.section.length))
	sum.table$finish.section.length<-as.numeric(as.character(sum.table$finish.section.length))
	sum.table$pathway.pass[is.na(sum.table$pathway)]<-1
	names(sum.table)[1:2]<-c("start","end")	
}

} # close if loop to check to see if ther are more than one barriers

else {

sum.table<-data.frame
# "sink" has to appear twic in the start column, and the second section has to appear twice in the end
 
start<-c(sections[1],sections[1],sections[2],sections[2])
end<-c(sections[1],sections[2],sections[1],sections[2])
path2<-c(sections[1],paste(sections,collapse=","),paste(rev(sections),collapse=","),sections[2])
barrier.id<-c("NA",passability$Bar_ID[1],passability$Bar_ID[1],"NA")
pathway.pass<-c(1,passability$Pass[1],passability$Pass[1],1)
start.section.length<-lengths$Shape_Length[match(start, lengths$Seg_ID)]
finish.section.length<-lengths$Shape_Length[match(end, lengths$Seg_ID)]
sum.table<-data.frame(start,end,path2,barrier.id,pathway.pass,start.section.length,finish.section.length)
}


return(sum.table)
		

}