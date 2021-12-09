# VBA-Script-to-concatenate-unique-cells

To delete the commas from each cell of the selection quickly and easily, you can apply the following formulas, please do as follows:

1. Enter this formula =IF(RIGHT(A2,1)=",",LEFT(A2,LEN(A2)-1),A2) 
