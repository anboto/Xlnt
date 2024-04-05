// SPDX-License-Identifier: Apache-2.0
// Copyright 2021 - 2024, the Anboto author and contributors
#include <Core/Core.h>
#include <Functions4U/Functions4U.h>
#include "Xlnt.h"

namespace Upp {

String GetCell(int col, int row) {
	ASSERT(row > 0 && col > 0);
    String colStr;
    while (col > 0) {
        colStr.Insert(0, static_cast<char>('A' + (col - 1) % 26));
        col = (col - 1) / 26;
    }
    return colStr + FormatInt(row);
}

void XlsxFill(xlnt::worksheet &ws, const Grid &g, bool removeEmpty) {
	for (int c = 0; c < g.cols(); ++c)
		ws.column_properties(c+1).width = g.GetWidth(c)/5.;
	
	if (g.GetNumHeaderRows() > 0) 
		ws.range(xlnt::range_reference(xlnt::cell_reference(1, 1),
                 					   xlnt::cell_reference(g.cols(), g.GetNumHeaderRows()))).font(xlnt::font().bold(true))
                 					   														 .fill(xlnt::fill::solid(xlnt::rgb_color(230, 230, 230)));
	
	if (g.GetNumHeaderCols() > 0) 
		ws.range(xlnt::range_reference(xlnt::cell_reference(1, 1),
                 					   xlnt::cell_reference(g.GetNumHeaderCols(), g.rows()))).font(xlnt::font().bold(true))
                 					   														 .fill(xlnt::fill::solid(xlnt::rgb_color(230, 230, 230)));
	
	for (int r = 0, row = 0; r < g.rows(); ++r) {
		bool printRow = true;
		if (removeEmpty) {			// Doesn't show empty rows
			printRow = false;
			for (int c = 0; c < g.cols(); ++c) {
				if (!IsNull(g.Get(r, c))) {
					printRow = true;	
					break;
				}
			}
		}
		if (printRow) {
			for (int c = 0; c < g.cols(); ++c) {
				const Value &v = g.Get(r, c);
				if (v.Is<bool>())
					ws.cell(c+1, row+1).value(bool(v));
				else if (v.Is<int>())
					ws.cell(c+1, row+1).value(int(v));
				else if (v.Is<double>())
					ws.cell(c+1, row+1).value(double(v));
				else 
					ws.cell(c+1, row+1).value(~v.ToString());
			}
			row++;
		}
	}	
}

}