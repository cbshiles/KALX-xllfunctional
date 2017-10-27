// curry.cpp - Curry a function
// Copyright (c) 2006-2010 KALX, LLC. All rights reserved. No warranty is made.
/*
f(x,y) = 2*x + y 
f=BIND(ADD, BIND(MUL,2,.),.)
CALL(f, 3, 4)
z=CALL(BIND(MUL,2,.), 3) -> 6
w=CALL(ADD, 6, 4) -> 10
*/
#include <functional>
#include "functional.h"

using namespace xll;
using namespace function;

typedef traits<XLOPERX>::xword xword;

static AddInX xai_functional_bind(
	FunctionX(XLL_HANDLEX XLL_UNCALCEDX, _T("?xll_functional_bind"), _T("FUNCTIONAL.BIND"))
	.Arg(XLL_HANDLEX, _T("Function"), _T("is the function to be curried."))
	.Arg(XLL_LPOPERX,  _T("Arg"), _T("is an array of arguments to be curried. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a handle to a function curried with the non-missing args."))
	.Documentation(
		_T("The arguments of the curried function are those that are missing. ")
		_T("Use FUNCTIONAL.CALL to supply these arguments and call the function.")
		,
		xml::xlink(_T("FUNCTIONAL.CALL"))
	)
);
HANDLEX WINAPI
xll_functional_bind(HANDLEX f, const LPOPER px)
{
#pragma XLLEXPORT
 	handle<bind> fx(new bind(f, *px));

	return fx.get();
}

static AddInX xai_functional_call(
	FunctionX(XLL_LPOPERX, _T("?xll_functional_call"), _T("FUNCTIONAL.CALL"))
	.Arg(XLL_HANDLEX, _T("Handle"), _T("is the handle of the curried function returned by FUNCTIONAL.BIND."))
	.Arg(XLL_LPOPERX, _T("Arg1"), _T("is an array of arguments that wer missing in the call to FUNCTIONAL.BIND. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Call a function by supplying the missing arguments to FUNCTIONAL.BIND."))
	.Documentation(
		_T("The arguments are those that were not supplied to FUNCTIONAL.BIND. ")
	)
);
LPOPER WINAPI
xll_functional_call(HANDLEX f, const LPOPER px)
{
#pragma XLLEXPORT
	static OPER o;

	try {

		handle<bind> fx_(f);
		if (fx_)
			o = fx_->call(*px);
		else
			o = bind(f, *px).call();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = Err(xlerrValue);
	}

	return &o;
}

// 16 LPOPERs
#define XLL_LPOPERSX XLL_LPOPERX XLL_LPOPERX XLL_LPOPERX XLL_LPOPERX XLL_LPOPERX XLL_LPOPERX XLL_LPOPERX XLL_LPOPERX \
                 XLL_LPOPERX XLL_LPOPERX XLL_LPOPERX XLL_LPOPERX XLL_LPOPERX XLL_LPOPERX XLL_LPOPERX XLL_LPOPERX
/*
static AddInX xai_functional_args(
	FunctionX(XLL_LPOPERSX, _T("?xll_functional_args"), _T("FUNCTIONAL.ARGS"))
	.Arg(XLL_LPOPERSX, _T("Arg1, ..."), _T("are a numeric arguments or NA() to indicate a missing argument."))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return an array for the second argument to FUNCTIONAL.BIND."))
	.Documentation(
		_T("Non numeric arguments will not be bound.")
	)
);
LPOPER WINAPI
xll_functional_args(LPOPERX px)
{
#pragma XLLEXPORT
	static OPER o;

	o.resize(0,0);

	for (LPOPERX* ppx = &px; (*ppx)->xltype != xltypeMissing && ppx - &px < 16; ++ppx) {
		o.push_back(**ppx);
	}

	o.resize(1, o.size());

	return &o;
}
*/
#if 0
#ifdef _DEBUG

static AddInX xai_test_curry(
	FunctionX(XLL_DOUBLEX, _T("?xll_test_curry"), _T("TEST.CURRY"))
	.Arg(XLL_DOUBLEX, _T("arg1"), _T(""))
	.Arg(XLL_DOUBLEX, _T("arg2"), _T(""))
);
double WINAPI
xll_test_curry(double a1, double a2)
{
#pragma XLLEXPORT

	return a1 + 10*a2;
}

double
add(double x, double y)
{
	return x + y;
}
double
add2(size_t n, const double *x)
{
	ensure (n == 2);

	return x[0] + x[1];
}

int
test_curry(void)
{
	OPERX b, c;

	try {
		b = ExcelX(xlfEvaluate, OPERX(_T("=FUNCTIONAL.BIND(TEST.CURRY,,2)")));
		ensure (b.xltype == xltypeNum);
		c = ExcelX(xlUDF, ExcelX(xlfEvaluate, OPERX(_T("FUNCTIONAL.CALL"))), b, OPERX(1));
		ensure (c.xltype == xltypeNum);
		ensure (c.val.num == 1 + 10*2);

		b = ExcelX(xlfEvaluate, OPERX(_T("=FUNCTIONAL.BIND(TEST.CURRY,3,)")));
		ensure (b.xltype == xltypeNum);
		c = ExcelX(xlUDF, ExcelX(xlfEvaluate, OPERX(_T("FUNCTIONAL.CALL"))), b, OPERX(4));
		ensure (c.xltype == xltypeNum);
		ensure (c.val.num == 3 + 10*4);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
//static Auto<OpenAfter> xoa_test_curry(test_curry);

#endif

static AddInX xai_function_bind(
	FunctionX(XLL_HANDLEX XLL_UNCALCEDX, _T("?xll_function_bind"), _T("FUNCTIONAL.BIND"))
	.Arg(XLL_DOUBLEX, _T("function"), _T("is the function to be curried"))
	.Arg(XLL_LPOPERX,  _T("arg1?"), _T("is the first argument to be curried, or missing if it is not to be curried"))
	.Arg(XLL_LPOPERX,  _T("arg2?"), _T("is the second argument to be curried, or missing if it is not to be curried"))
	.Arg(XLL_LPOPERX,  _T("arg3?"), _T("is the third argument to be curried, or missing if it is not to be curried"))
	.Arg(XLL_LPOPERX,  _T("arg4?"), _T("is the fourth argument to be curried, or missing if it is not to be curried. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a handle to a function curried with the non-missing args."))
	.Documentation(
		_T("The arguments of the curried function are those that are missing. ")
		_T("Use FUNCTIONAL.CALL to supply these arguments and call the function.")
		,
		xml::xlink(_T("FUNCTIONAL.CALL"))
	)
);
HANDLEX WINAPI
xll_function_bind(double f, LPOPERX pa1, LPOPERX pa2, LPOPERX pa3, LPOPERX pa4)
{
#pragma XLLEXPORT
	try {
		OPERX x(4, 1);

		x[0] = *pa1;
		x[1] = *pa2;
		x[2] = *pa3;
		x[3] = *pa4;

		// func finds the arity of f
		handle<function::func> h = new function::func(f, x);

		return h.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return 0;
}

static AddInX xai_function_call(
	FunctionX(XLL_LPOPERX, _T("?xll_function_call"), _T("FUNCTIONAL.CALL"))
	.Arg(XLL_HANDLEX, _T("handle"), _T("is the handle of the curried function returned by FUNCTIONAL.BIND"))
	.Arg(XLL_LPOPERX, _T("arg1?"), _T("is the first argument that was missing in the call to FUNCTIONAL.BIND"))
	.Arg(XLL_LPOPERX, _T("arg2?"), _T("is the second argument that was missing in the call to FUNCTIONAL.BIND"))
	.Arg(XLL_LPOPERX, _T("arg3?"), _T("is the third argument that was missing in the call to FUNCTIONAL.BIND"))
	.Arg(XLL_LPOPERX, _T("arg4?"), _T("is the fourth argument that was missing in the call to FUNCTIONAL.BIND. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Call a function by supplying the missing arguments to FUNCTIONAL.BIND."))
	.Documentation(
		_T("The arguments are those that were not supplied to FUNCTIONAL.BIND. ")
		_T("If <codeInline>Handle</codeInline> is not a handle, then that is returned and the ")
		_T("arguments are ignored. ")
	)
);
LPOPERX WINAPI
xll_function_call(HANDLEX h, LPOPERX pa1, LPOPERX pa2, LPOPERX pa3, LPOPERX pa4)
{
#pragma XLLEXPORT
	static OPERX o(ErrX(xlerrValue));

	try {
		handle<function::func> h_(h);

		if (h_) {
			if (pa1->xltype == xltypeMissing) {
				o = h_->call();
			}
			else if (pa2->xltype == xltypeMissing) {
				o = h_->call(*pa1);
			}
			else if (pa3->xltype == xltypeMissing) {
				o = h_->call(*pa1, *pa2);
			}
			else if (pa4->xltype == xltypeMissing) {
				o = h_->call(*pa1, *pa2, *pa3);
			}
			else {
				o = h_->call(*pa1, *pa2, *pa3, *pa4);
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}
#endif