// realtime.cpp - update cells in realtime
// Copyright (c) 2006-2010 KALX, LLC. All rights reserved. No warranty is made.
#include "functional.h"

using namespace xll;

static AddInX xai_realtime(
	FunctionX(XLL_LPOPERX XLL_UNCALCEDX, _T("?xll_realtime"), _T("FUNCTIONAL.REALTIME"))
	.Arg(XLL_DOUBLEX, _T("function"), _T("is a user defined function"))
	.Arg(XLL_DOUBLEX, _T("cell"), _T("is the cell containing a value"))
	.Arg(XLL_DOUBLEX, _T("init"), _T("is the initial value to be used. "))
	.FunctionHelp(_T("Apply function to Cell and its previous value."))
	.Category(CATEGORY)
	.Documentation(
		_T("When a cell is recalulated using F2-Enter, the result is the ")
		_T("<codeInline>function</codeInline> applied to the current ")
		_T("<codeInline>cell</codeInline> value and <codeInline>init</codeInline>. ")
	)
);
LPOPERX WINAPI
xll_realtime(double f, double x, double x0)
{
#pragma XLLEXPORT
	static OPERX y;

	try {
		LOPERX o = ExcelX(xlCoerce, ExcelX(xlfCaller));

		if (o.xltype == xltypeNum) {
			if (o.val.num == 0)
				o.val.num = x0;
			y = Excel<XLOPERX>(xlUDF, OPERX(f), OPERX(x), OPERX(o.val.num));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		if (y.xltype != xltypeErr)
			y = Err(xlerrValue);
	}

	return &y;
}