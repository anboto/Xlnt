// SPDX-License-Identifier: Apache-2.0
// Copyright 2021 - 2024, the Anboto author and contributors
#include <Core/Core.h>

using namespace Upp;

#include <Xlnt/Xlnt.h>


CONSOLE_APP_MAIN
{
	StdLogSetup(LOG_COUT|LOG_FILE);
	
	try {
		{	// Creating and saving
			Vector<Vector<double>> data;
		    for (int r = 0; r < 100; r++) {
		        
		        Vector<double> &row = data.Add();
				for(int c = 0; c < 2; c++)
					row << c+1;
		    }
		    Vector<String> header = {"First", "Second"};
		    
		    xlnt::workbook wb;
		    xlnt::worksheet ws = wb.active_sheet();		//Creating the output worksheet
		    ws.title("My Data");
		    
		    ws.cell("A1").value(~header[0]);
		    ws.cell("B1").value(~header[1]);
		    ws.cell("C1").value("Sum");
		    ws.range("A1:C1").font(xlnt::font().bold(true));
		    
		    for (int r = 0; r < data.size(); r++) {
		        const Vector<double> &row = data[r];
		        for (int c = 0; c < row.size(); c++)
				    ws.cell(xlnt::cell_reference(c + 1, r + 2)).value(row[c]);
		        ws.cell(xlnt::cell_reference(row.size() + 1, r + 2)).formula(~Format("=SUM(%s:%s)", GetCell(1, r+2), GetCell(row.size(), r+2)));
		    }
		    wb.save("output.xlsx"); 
		}
		{	// Loading
		    xlnt::workbook wb;
    		wb.load("output.xlsx");
    		xlnt::worksheet ws = wb.active_sheet();
    		for (auto row : ws.rows(false)) 
        		for (auto cell : row) {
        			if (cell.has_formula())
        				UppLog() << "\nFormula: " << cell.formula();
					UppLog() << "\nValue: " << cell.to_string();
        		}
		}
	} catch (Exc err) {
		UppLog() << "\n" << Format(t_("Problem found: %s"), err);
		SetExitCode(-1);
	} catch (...) {
		UppLog() << "\n" << t_("Problem found");
		SetExitCode(-1);
	}
		
	UppLog() << "\nProgram ended\n";
	#ifdef flagDEBUG
	ReadStdIn();
	#endif
}

