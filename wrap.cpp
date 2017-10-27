// wrap.cpp - wrapped up arguments
#include "functional.h"

using namespace xll;

typedef traits<XLOPERX>::xcstr xcstr;

static AddInX xai_function_wrap(
	FunctionX(XLL_LPOPERX, _T("?xll_function_wrap"), _T("FUNCTIONAL.WRAP"))
	.Arg(XLL_DOUBLEX, _T("function"), _T("is a user-defined function."))
	.Arg(XLL_LPOPERX, _T("array"), _T("is an array of at most four elements to be used as arguments to the user-defined function. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Call function with arguments that have been wrapped up in an array."))
	.Documentation(
		_T("Some add-in functions return arrays whose elements need to be used as individual arguments. ")
		_T("For example, <codeInline>FUNCTIONAL.WRAP(FUNCTION.SUM, {2,3})</codeInline> returns ")
		_T("the same value as <codeInline>FUNCTION.SUM(2,3)</codeInline>. ")
	)
);
LPOPERX WINAPI
xll_function_wrap(double f, LPOPERX po)
{
#pragma XLLEXPORT
	static OPERX oResult;

	try {
		if (po->xltype == xltypeMissing) {
			oResult = ExcelX(xlUDF, OPERX(f));
		}
		else if (po->xltype != xltypeMulti) {
			oResult = ExcelX(xlUDF, OPERX(f), *po);
		}
		else {
			switch (po->size()) {
			case 1:
				oResult = ExcelX(xlUDF, OPERX(f), (*po)[0]);
				break;
			case 2:
				oResult = ExcelX(xlUDF, OPERX(f), (*po)[0], (*po)[1]);
				break;
			case 3:
				oResult = ExcelX(xlUDF, OPERX(f), (*po)[0], (*po)[1], (*po)[2]);
				break;
			case 4:
				oResult = ExcelX(xlUDF, OPERX(f), (*po)[0], (*po)[1], (*po)[2], (*po)[3]);
				break;
			default:
				oResult = ErrX(xlerrValue);
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		oResult = ErrX(xlerrValue);
	}

	return &oResult;
}