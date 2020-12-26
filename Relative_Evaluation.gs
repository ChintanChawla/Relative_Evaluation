var spreadSheet=SpreadsheetApp.getActiveSpreadsheet();
var result_sheet=spreadSheet.getSheetByName("relative_improvement")
var data_len=result_sheet.getDataRange().getValues().length-4
var number_sheets=result_sheet.getRange("G3").getValue();
function onOpen()
 {
   
       SpreadsheetApp.getUi()
       .createMenu('Script for evaluation')
       .addItem('Run','main')
       .addItem('Add Templete Sheet','addSheets')
       .addToUi()
 
 } 
/*
 SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Script for evaluaion')
      .addItem('Run', 'main')
      .addToUi();
      */
   
function main() 
{  
   
   // get if the sheet is validated. Use validate function to return boolean
   // If not, get the error response from the validate funciton.
   // Show the error respone in an alert object and stop execution there. 
   if (validate()==false)
      {
         SpreadsheetApp.getUi().alert("Please Check the range of sheets to be evaluated.Range is required to between 2 and TOTAL NUMBER of result sheets");
         return
      };
   var date=Utilities.formatDate(new Date(), "GMT+1", "MM/dd/yyyy");
   
   ;
   result_sheet.getRange("G2").setValue(date)//setting date
   // Writting relative improvement on sheet 
   for(var i=1;i<=data_len;i++)
    {
      
     var name=result_sheet.getRange(i+4,1,1, 1).getValue()
     if(name=="")
     {
       result_sheet.getRange(i+4,3,1,1).setValue(-3000);
     }
     else
     {
       result_sheet.getRange(i+4,3,1,1).setValue(relative_inc(name));
     }
      
    } 
   sort_data();
   rank();
   for(var k=1;k<=data_len;k++)
   {

     var cell_check_range=result_sheet.getRange(k+4,3,1,1)
     var cell_check=cell_check_range.getValue()
     if (cell_check==-2000)
     {
      cell_check_range.setValue("No previous data")
     }
     if(cell_check==-3000)
     {
      cell_check_range.setValue("Unauthorized Student")
     }
   
   }

}


// This function validates the number of sheets to be evaluated
// Input and output of this function is null
function validate()
   {
     
    var sheets_num=spreadSheet.getSheets().length-1;
    Logger.log(sheets_num); 
    var requested_result_range = SpreadsheetApp.getActive().getSheetByName("relative_improvement").getRange("G3").getValue()
    if(sheets_num < requested_result_range || requested_result_range==1) 
    {
      return false; 
    }
    else 
    {
     return true;
    }  
    
   }


/** This function is used to import copy of result sheet 
* There is no input or output
**/
function addSheets()
 {
  
  var source = SpreadsheetApp.openById('188MO4olB_VU8sWh8t_ryXXCB1jhvISpIjYPmNp9SNPo');
  var sourcename = source.getSheetName();
  var sourceSheet=source.getSheets()[0];

  var destination = spreadSheet;
  sourceSheet.copyTo(destination);
  var sheet = destination.getSheetByName("copy of result1");
  sheet.setName("result_NO") 

 }


/* This function gives the appropriate previous marks
 * It take name as input
 * return the max marks out of all previous tets,
 */
 function prev_marks(name)
 {
  
  var max_marks=0;
  for(var i=1;i<number_sheets;i++)
   {
     
     var sheet_data=spreadSheet.getSheetByName("result"+i).getDataRange().getValues();
     var marks=get_marks(sheet_data,name)/get_goal(sheet_data);
     if(marks>max_marks)
      {      
       max_marks=marks;
      }
      
   }
   
   if(max_marks==0)
   {
     return -2000;
   }
   
 
    return max_marks;
  
  }
/** This funtion returns the marks of the given studtent 
 *First parameter takes the data of sheet
 *Second parameter is the name of the student
**/
 function get_marks(data,name)
 {
 
  for (var j=0;j<data.length;j++)
   {
    if(data[j][0]==name)
     {
      return data[j][data[j].length-3];
     }
    
  }
 
 }
/** 
 *This function returns the current value of marks
 *It takes name as parameter
 **/

 function current_marks(name)
 {
   
   var sheet_data=spreadSheet.getSheetByName("result"+number_sheets).getDataRange().getValues()
   
   return  get_marks(sheet_data,name)/get_goal(sheet_data);
 
 
 }
 
// This funcrion returns the difference between total marks of last 2 tests
function diff(marks1,marks2) 
{
         
         var diff=marks2-marks1;
         return diff; 
}


// this is to get the goal



/*
  *This function is used to set goal for the marks
  * It takes sheet as input
  * It returns the final goal*/

function get_goal(sheet)
{
 
  return sheet[2][sheet[2].length-2]
  
 }
   
   
   
  
/**
 *This funtion calculates the relative increment
 *Parameter 1 are total marks of test 1
 *Parameter 2 are total marks of test1
 It returns the % of improvement**/
function relative_inc(name)
{
  
 var marks1=prev_marks(name);
 var marks2=current_marks(name);
 var improvement=diff(marks1,marks2)/(1-marks1)*100;
 if(marks1==-2000)
 {
  return -2000;
 }
 return improvement;
 
 }
/**
 * This function sorts the studens according to their difference in marks 
 * it returns null
**/ 
function sort_data() 
  {
   var sheet_data=result_sheet.getDataRange().getValues();
   var range= result_sheet.getRange(5,1,sheet_data.length-4,sheet_data[3].length-1);
   range.sort({column: 3, ascending: false});
  
  }
  
  //This funtion rank the sudents according to improvement in marks 
  function rank()
  {
  
   for(var i=1;i<=data_len;i++)
   {
     result_sheet.getRange(i+4,4,1,1).setValue(i);
     if(i+4>5 && result_sheet.getRange(i+4,3,1,1).getValue()==result_sheet.getRange(i+3,3,1,1).getValue())
      {
        result_sheet.getRange(i+4,4,1,1).setValue(result_sheet.getRange(i+3,4,1,1).getValue());
      }
   }
  
  }

