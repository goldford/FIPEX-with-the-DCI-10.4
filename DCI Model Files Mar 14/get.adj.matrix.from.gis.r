get.adj.matrix.from.gis<-function(segment.matrix=read.csv("segment_matrix.csv")){


#### CREATE ADJACENCY MATRIX FROM ARCGIS OUTPUT ####
#output: a matrix with 1's and 0's that describe whether or not segments neighbour eachother

#The input is the segment matrix, which is really two vectors. For each segment in the first column (Seg_ID), the second column (Seg) gives the id of other sections it touches, including itself.

# obtain the segments in whatever order they are in the file
segments<-with(segment.matrix, unique(Seg_ID))

#create a matrix with only 0's in them
adj.matrix<-matrix(nrow=length(segments), ncol=length(segments), rep(0,length(segments)*length(segments)))

segment.length<-length(segments)

rownames(adj.matrix)<-colnames(adj.matrix)<-segments

for (i in 1:segment.length)
	{
	# find the segments in segment.matrix$Seg where segment.matrix$Seg_ID matches segments[i]
	# index of matching positions - this will be a vector of 1's and NA's
	pos.match<-	match(segment.matrix$Seg_ID,segments[i])
	# keep only the positions where pos.match==1
	adj.segments<-segment.matrix$Seg[!is.na(pos.match)]

	# find the column positions that correspond to the adjacaent segments
	col<-match(adj.segments,segments)

	# the row number should correspond to i
	row<-i

	# assign a value of 1 for all values of row and col
	adj.matrix[row,col]<-1
	}

write.table(adj.matrix,"adjacency matrix.csv",row.names=F, sep=",")
#write.table seems to give problems when reading in "adjacency matrix.csv" - it doesn't recognize that it's a matrix so it creates column rows with headings "X1,X2,X3,..." instead of "1,2,3,..."

#in order to avoid the problem above return the object:
return(adj.matrix)


}