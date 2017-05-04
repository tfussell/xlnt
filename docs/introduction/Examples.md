## Examples
### Simple - reading from an existing xlsx spread sheet.
The following C plus plus code will read the values from an xlsx file and print the string values to the screen. This is a very simple example to get you started.

```
#include <iostream>
#include <xlnt/xlnt.hpp>

int main()
{
    xlnt::workbook wb;
    wb.load("/home/timothymccallum/test.xlsx");
    auto ws = wb.active_sheet();
    std::clog << "Processing spread sheet" << std::endl;
    for (auto row : ws.rows(false)) 
    { 
        for (auto cell : row) 
	{ 
	    std::clog << cell.to_string() << std::endl;
	}
    }
    std::clog << "Processing complete" << std::endl;
    return 0;
}
```
Save the contents of the above file 
```
/home/timothymccallum/process.cpp
```
Compile by typing the following command
```
g++ -std=c++14 -lxlnt process.cpp -o process
```
Excecute by typing the following command
```
./process
```
The output of the program, in my case, is as follows
```
Processing spread sheet
This is cell A1.
This is cell B1
… and this is cell C1
We are now on the second row at cell A2
B2
C2
Processing complete
```
As you can see the process.cpp file simply walks through the spread sheet values row by row and column by column (A1, B1, C1, A2, B2, C2 and so on).

### Simple - storing a spread sheet in a 2 dimensional C++ Vector for further processing
Loading a spread sheet into a Vector provides oppourtunities for you to perform high performance processing. There will be more examples on performing fast look-ups, merging data, performing deduplication and more. For now, let's just learn how to get the spread sheet loaded into memory.
```
#include <iostream>
#include <xlnt/xlnt.hpp>
#include <vector>

int main()
{
    xlnt::workbook wb;
    wb.load("/home/timothymccallum/test.xlsx");
    auto ws = wb.active_sheet();
    std::clog << "Processing spread sheet" << std::endl;
    std::clog << "Creating a single vector which stores the whole spread sheet" << std::endl;
    std::vector< std::vector<std::string> > theWholeSpreadSheet;
    for (auto row : ws.rows(false)) 
    { 
        std::clog << "Creating a fresh vector for just this row in the spread sheet" << std::endl;
	std::vector<std::string> aSingleRow;
	for (auto cell : row) 
	{ 
	    std::clog << "Adding this cell to the row" << std::endl;
	    aSingleRow.push_back(cell.to_string());
	}
	std::clog << "Adding this entire row to the vector which stores the whole spread sheet" << std::endl;
	theWholeSpreadSheet.push_back(aSingleRow);
    }
    std::clog << "Processing complete" << std::endl;
    std::clog << "Reading the vector and printing output to the screen" << std::endl;
    for (int rowInt = 0; rowInt < theWholeSpreadSheet.size(); rowInt++)
    {
        for (int colInt = 0; colInt < theWholeSpreadSheet.at(rowInt).size(); colInt++)
	{
	    std::cout << theWholeSpreadSheet.at(rowInt).at(colInt) << std::endl;
        }
    }
    return 0;
}
```
Save the contents of the above file 
```
/home/timothymccallum/process.cpp
```
Compile by typing the following command
```
g++ -std=c++14 -lxlnt process.cpp -o process
```
Excecute by typing the following command
```
./process
```
The output of the program, in my case, is as follows
```
Processing spread sheet
Creating a single vector which stores the whole spread sheet
Creating a fresh vector for just this row in the spread sheet
Adding this cell to the row
Adding this cell to the row
Adding this cell to the row
Adding this entire row to the vector which stores the whole spread sheet
Creating a fresh vector for just this row in the spread sheet
Adding this cell to the row
Adding this cell to the row
Adding this cell to the row
Adding this entire row to the vector which stores the whole spread sheet
Processing complete
Reading the vector and printing output to the screen
This is cell A1.
This is cell B1
… and this is cell C1
We are now on the second row at cell A2
B2
C2
```
You will have noticed that this process is very fast. If you type the "time" as shown below, you can measure just how fast loading and retrieving your spread sheet is, using xlnt; In this case only a fraction of a second. More on this later.
```
time ./process 
...
real	0m0.044s
```
### Simple - writing values to a new xlsx spread sheet.
```
#include <iostream>
#include <xlnt/xlnt.hpp>
#include <vector>
#include <string>

int main()
{
    //Creating a 2 dimensional vector which we will write values to
    std::vector< std::vector<std::string> > wholeWorksheet;
    //Looping through each row (100 rows as per the second argument in the for loop)
    for (int outer = 0; outer < 100; outer++)
    {
        //Creating a fresh vector for a fresh row
	std::vector<std::string> singleRow;
	//Looping through each of the columns (100 as per the second argument in the for loop) in this particular row
	for(int inner = 0; inner < 100; inner++)
	{
	    //Adding a single value in each cell of the row 
	    std::string val = std::to_string(inner + 1);
	    singleRow.push_back(val);			
	}
	//Adding the single row to the 2 dimensional vector
	wholeWorksheet.push_back(singleRow);
	std::clog << "Writing to row " << outer << " in the vector " << std::endl;
    }
    //Writing to the spread sheet
    //Creating the output workbook 
    std::clog << "Creating workbook" << std::endl;
    xlnt::workbook wbOut;
    //Setting the destination output file name
    std::string dest_filename = "output.xlsx";
    //Creating the output worksheet
    xlnt::worksheet wsOut = wbOut.active_sheet();
    //Giving the output worksheet a title/name
    wsOut.title("data");
    //We will now be looping through the 2 dimensional vector which we created above
    //In this case we have two iterators one for the outer loop (row) and one for the inner loop (column)
    std::clog << "Looping through vector and writing to spread sheet" << std::endl;
    for (int fOut = 0; fOut < wholeWorksheet.size(); fOut++)
    {
        std::clog << "Row" << fOut << std::endl;
        for (int fIn = 0; fIn < wholeWorksheet.at(fOut).size(); fIn++)
        {
            //Take notice of the difference between accessing the vector and accessing the work sheet
	    //As you may already know Excel spread sheets start at row 1 and column 1 (not row 0 and column 0 like you would expect from a C++ vector) 
	    //In short the xlnt cell reference starts at column 1 row 1 (hence the + 1s below) and the vector reference starts at row 0 and column 0
	    wsOut.cell(xlnt::cell_reference(fIn + 1, fOut + 1)).value(wholeWorksheet.at(fOut).at(fIn));
	    //Further clarification to avoid confusion
	    //Cell reference arguments are (column number, row number); e.g. cell_reference(fIn + 1, fOut + 1)
	    //Vector arguments are (row number, column number); e.g. wholeWorksheet.at(fOut).at(fIn)
	}
    }
    std::clog << "Finished writing spread sheet" << std::endl;
    wbOut.save(dest_filename); 
    return 0;
}
```
This process is also quite quick; a time command showed that xlnt was able to create and write 10, 000 values to the output spread sheet in 0.582 seconds.
