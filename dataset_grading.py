import pandas as pd 
import math
import os

THIS_FOLDER = os.path.dirname(os.path.abspath(__file__))

actual_data_set_location= os.path.join(THIS_FOLDER, "actual_data_set.xlsx")
test_data_set_location= os.path.join(THIS_FOLDER,"test.xlsx")
output_data_set_location= os.path.join(THIS_FOLDER,"output.xlsx")


#reading the actual data set into a pandas dataframe
pd_data_set=pd.read_excel(actual_data_set_location)
lst_data_set = list(pd_data_set.values)

#reading the test dataset into a pandas dataframe
pd_test_data=pd.read_excel(test_data_set_location)
lst_test_data= list(pd_test_data.values[0])[1:]

#putting the column names into a list to prepare for the output
lst_column = list(pd_data_set.columns.values)
#inserting the MATCH_SCORE column name
lst_column.insert(1,"MATCH_SCORE")
#creating a list with empty dataset but only columns
lst_out_put = [lst_column]

#we get all the rows from data_set.values
for data_row in pd_data_set.values:
    #match_score holds the value assigned to the actual data row after matching with the test data
    match_score=0
    index = 0
    print("Scoring variable set : "+str(data_row[0]))
    #data[0] holds the variable name (i.e. A,B,C,D...) so skipping 0th index and taking rest using data_row[1:]
    for data in data_row[1:]:
        print("matching "+str(data)+" with "+str(lst_test_data[index])+" at index "+str(index))
        match_score += (data - lst_test_data[index])**2
        index += 1
    #applying square root average to calculate match score
    match_score = 1- (math.sqrt(match_score)/11) 
    data_row = list(data_row) 
    #inserting the match score after variable name
    data_row.insert(1,match_score)
    lst_out_put.append(data_row)


#writing output into the output excel after wrapping it into a pandas object
pd_out_put = pd.DataFrame(lst_out_put)
pd_out_put.to_excel(output_data_set_location)
print("Analysis complete: Please open "+output_data_set_location+" to check the output.")
