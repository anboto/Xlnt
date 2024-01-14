// SPDX-License-Identifier: Apache-2.0
// Copyright 2021 - 2024, the Anboto author and contributors
#include <Core/Core.h>

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

}