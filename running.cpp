// running.cpp - running binary operation of elements of an array
// Copyright (c) 2006-2009 KALX, LLC. All rights reserved. No warranty is made.
#pragma warning(disable: 4996)
#include <functional>
#include <numeric>
#include "array.h"

using namespace array;
using namespace xll;

typedef traits<XLOPERX>::xcstr xcstr;
typedef traits<XLOPERX>::xfp xfp;
typedef traits<XLOPERX>::xword xword;

static AddInX xai_running(
	FunctionX(XLL_FPX, _T("?xll_running"), _T("ARRAY.RUNNING"))
	.Arg(XLL_FPX, _T("Array"), _T("is an array of numbers. "))
	.Arg(XLL_DOUBLEX, _T("Function"), _T("is a user defined function taking two arguments. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns the running binary Function applied to the elements of Array."))
	.Documentation(
		_T("For example, <codeInline>ARRAY.RUNNING({2, 3, 4}, FUNCTION.MUL)</codeInline> ")
		_T("returns <codeInline>{2, 6, 24}</codeInline>. ")
		_T("More generally, <codeInline>ARRAY.RUNNING({a, b, c}, f)</codeInline> ")
		_T("returns <codeInline>{a, f(a, b), f(f(a, b), c)}</codeInline>.")
	)
);
xfp* WINAPI
xll_running(xfp* pa, double f)
{
#pragma XLLEXPORT
	try {
		if (FPX* px = array::handle(pa))
			pa = &(*px);

		if (f)
			std::partial_sum(pa->array, pa->array + size(*pa), &pa->array[0], binop(f)); 
		else
			std::partial_sum(pa->array, pa->array + size(*pa), &pa->array[0]); 
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return pa;
}

#ifdef _DEBUG
#ifndef EXCEL12

static AddInX xai_test_running(
	_T("?xll_test_running"), XLL_DOUBLEX XLL_DOUBLEX XLL_DOUBLEX,
	_T("TEST.RUNNING"), _T("Arg1, Arg2")
);
double WINAPI
xll_test_running(double a1, double a2)
{
#pragma XLLEXPORT

	return a1 + a1*a2;
}

int
test_running(void)
{
	try {
		OPER o;

		o = ExcelX(xlUDF
			, OPERX(_T("ARRAY.RUNNING"))
			, ExcelX(xlfEvaluate, OPERX(_T("{1,1,1}")))
			, ExcelX(xlfEvaluate, OPERX(_T("=TEST.RUNNING")))
		);
		ensure (o.xltype == xltypeMulti);
		ensure (o.size() == 3);
		ensure (o[0] == 1);
		ensure (o[1] == 2);
		ensure (o[2] == 4);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<OpenAfter> xao_test_running(test_running);

#endif // EXCEL12
#endif // _DEBUG