// apply.cpp - like ARRAY.APPLY except for curried functions
// Copyright (c) 2006-2009 KALX, LLC. All rights reserved. No warranty is made.
#pragma warning(disable: 4127)
#include "functional.h"

using namespace xll;

// from xllarray/array.h
namespace array {
	inline xll::FPX*
	handle(const xfp* px)
	{
		if (xll::size(*px) > 1)
			return 0;
	
		xll::handle<xll::FPX> h((HANDLEX)px->array[0], true);
	
		return h;
	}
}

// call curried function or UDF
// nullary
template<class F> 
inline OPERX Call(F f);
template<> 
inline OPERX Call<double>(double f) 
{
	return ExcelX(xlUDF, OPERX(f));
}
template<> 
inline OPERX Call<xll::handle<function::bind>>(xll::handle<function::bind> f)
{
	return f->call();
}

template<class F>
xfp*
xll_apply0(F f, xword r, xword c)
{
	static FPX x;

	if (c == 0)
		c = 1;
	x.reshape(r, c);
	for (xword i = 0; i < x.size(); ++i)
		x[i] = Call(f);

	return x.get();
}

static AddInX xai_functional_apply0(
	FunctionX(XLL_FPX, _T("?xll_functional_apply0"), _T("FUNCTIONAL.APPLY0"))
	.Arg(XLL_DOUBLEX, _T("function"), IS_FUNCTIONAL)
	.Arg(XLL_WORDX, _T("rows"), IS_ROWS)
	.Arg(XLL_WORDX, _T("columns?"), IS_COLUMNS)
	.Category(CATEGORY)
	.FunctionHelp(_T("Return array with specified rows and columns filled with calls to function"))
	.Documentation(
		_T("Typically <codeInline>function</codeInline> is volatile. ")
	)
);
xfp* WINAPI
xll_functional_apply0(double f, xword r, xword c)
{
#pragma XLLEXPORT
	xfp* px;

	try {
		handle<function::bind> f_(f);
		ensure (f_);

		px = xll_apply0(f_, r, c);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return px;
}
static AddInX xai_udf_apply0(
	FunctionX(XLL_FPX, _T("?xll_udf_apply0"), _T("UDF.APPLY0"))
	.Arg(XLL_DOUBLEX, _T("function"), IS_UDF)
	.Arg(XLL_WORDX, _T("rows"), IS_ROWS)
	.Arg(XLL_WORDX, _T("columns?"), IS_COLUMNS)
	.Category(CATEGORY)
	.FunctionHelp(_T("Return array with specified rows and columns filled with calls to function"))
	.Documentation(
		_T("Typically <codeInline>function</codeInline> is volatile. ")
	)
);
xfp* WINAPI
xll_udf_apply0(double f, xword r, xword c)
{
#pragma XLLEXPORT
	xfp* px;

	try {
		px = xll_apply0(f, r, c);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return px;
}

// unary
template<class F> 
inline OPERX Call(F f, const OPERX& x);
template<> 
inline OPERX Call<double>(double f, const OPERX& x) 
{
	return ExcelX(xlUDF, OPERX(f), x);
}
template<> 
inline OPERX Call<xll::handle<function::bind>>(xll::handle<function::bind> f, const OPERX& x)
{
	return f->call(x);
}
template<class F>
xfp*
xll_apply1(F f, xfp* pa)
{
	static FPX x;

	FPX* px = array::handle(pa);
	if (px) {
		for (xword i = 0; i < px->size(); ++i)
			px[i] = Call(f, OPERX((*px)[i]));
	}
	else {
		for (xword i = 0; i < size(*pa); ++i)
			pa->array[i] = Call(f, OPERX(pa->array[i]));
	}

	return pa;
}

static AddInX xai_functional_apply1(
	FunctionX(XLL_FPX, _T("?xll_functional_apply1"), _T("FUNCTIONAL.APPLY1"))
	.Arg(XLL_DOUBLEX, _T("function"), IS_FUNCTIONAL)
	.Arg(XLL_FPX, _T("array"), IS_ARRAY_OR_HANDLE _T(". "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Applies function to each element of array"))
	.Documentation(
		_T("If <codeInline>array</codeInline> is a handle, the underlying array is modified. ")
	)
);
xfp* WINAPI
xll_functional_apply1(double f, xfp* pa)
{
#pragma XLLEXPORT
	try {
		handle<function::bind> f_(f);
		ensure (f_);

		pa = xll_apply1(f_, pa);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return pa;
}
static AddInX xai_udf_apply1(
	FunctionX(XLL_FPX, _T("?xll_udf_apply1"), _T("UDF.APPLY1"))
	.Arg(XLL_DOUBLEX, _T("Function"), IS_UDF)
	.Arg(XLL_FPX, _T("Array"), IS_ARRAY_OR_HANDLE _T(". "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Applies function to each element of array"))
	.Documentation(
		_T("If <codeInline>array</codeInline> is a handle, the underlying array is modified. ")
	)
);
xfp* WINAPI
xll_udf_apply1(double f, xfp* pa)
{
#pragma XLLEXPORT
	try {
		pa = xll_apply1(f, pa);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return pa;
}

// binary
template<class F> 
inline OPERX Call(F f, double x, double y);
template<> 
inline OPERX Call<double>(double f, double x, double y) 
{
	return ExcelX(xlUDF, OPERX(f), OPERX(x), OPERX(y));
}
template<> 
inline OPERX Call<xll::handle<function::bind>>(xll::handle<function::bind> f, double x, double y)
{
	OPERX arg(1, 2);

	arg[0] = x;
	arg[1] = y;

	return f->call(arg);
}
template<class F>
xfp*
xll_apply2(F f, xfp* pa1, xfp* pa2, BOOL outer, BOOL trans)
{
	static FPX x;

	FPX* px = array::handle(pa1);
	if (px) {
		for (xword i = 0; i < px->size(); ++i)
			px[i] = Call(f, (*px)[i], index(*pa2, i));
	}
	else {
		if (outer) {
			if (trans) {
				x.reshape(size(*pa2), size(*pa1));
				for (xword i = 0; i < x.columns(); ++i)
					for (xword j = 0; j < x.rows(); ++j)
						x(j, i) = Call(f, pa1->array[i], pa2->array[j]);
				pa1 = x.get();
			}
			else {
				x.reshape(size(*pa1), size(*pa2));
				for (xword i = 0; i < x.rows(); ++i)
					for (xword j = 0; j < x.columns(); ++j)
						x(i, j) = Call(f, pa1->array[i], pa2->array[j]);
				pa1 = x.get();
			}
		}
		else {
			for (xword i = 0; i < size(*pa1); ++i)
				pa1->array[i] = Call(f, pa1->array[i], index(*pa2, i));
		}
	}

	return pa1;
}
static AddInX xai_functional_apply2(
	FunctionX(XLL_FPX, _T("?xll_functional_apply2"), _T("FUNCTIONAL.APPLY2"))
	.Arg(XLL_DOUBLEX, _T("function"), IS_FUNCTIONAL)
	.Arg(XLL_FPX, _T("array1"), IS_ARRAY_OR_HANDLE)
	.Arg(XLL_FPX, _T("array2"), IS_ARRAY)
	.Arg(XLL_BOOLX, _T("outer?"), _T("is an optional boolean indicating the cartesian product is to be returned"))
	.Arg(XLL_BOOLX, _T("trans?"), _T("is an optional boolean indicating the transposed array is to be returned. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Applies function to each pair from array1 and array2"))
	.Documentation(
		_T("The returned array is the same shape as <codeInline>array1</codeInline>. ")
		_T("If <codeInline>array2</codeInline> is not the same shape as <codeInline>array1</codeInline> ")
		_T("then cyclic indexing is used. ")
		_T("If <codeInline>outer</codeInline> is true then a two dimensional array with rows equal to ")
		_T("the number of elements of <codeInline>array1</codeInline> and columns equal to ")
		_T("the number of elements of <codeInline>array2</codeInline> is returned with the ")
		_T("result of <codeInline>function</codeInline> called for each pair in the cartesian product ")
		_T("of the arrays. ")
		_T("If <codeInline>array1</codeInline> is a handle, the underlying array is modified. ")
	)
);
xfp* WINAPI
xll_functional_apply2(double f, xfp* pa1, xfp* pa2, BOOL outer, BOOL trans)
{
#pragma XLLEXPORT
	xfp* px;

	try {
		handle<function::bind> f_(f);
		ensure (f_);

		px = xll_apply2(f_, pa1, pa2, outer, trans);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return px;
}
static AddInX xai_udf_apply2(
	FunctionX(XLL_FPX, _T("?xll_udf_apply2"), _T("UDF.APPLY2"))
	.Arg(XLL_DOUBLEX, _T("function"), IS_UDF)
	.Arg(XLL_FPX, _T("array1"), IS_ARRAY_OR_HANDLE)
	.Arg(XLL_FPX, _T("array2"), IS_ARRAY)
	.Arg(XLL_BOOLX, _T("outer?"), _T("is an optional boolean indicating the cartesian product is to be returned"))
	.Arg(XLL_BOOLX, _T("trans?"), _T("is an optional boolean indicating the transposed array is to be returned. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Applies function to each pair from array1 and array2"))
	.Documentation(
		_T("The returned array is the same shape as <codeInline>array1</codeInline>. ")
		_T("If <codeInline>array2</codeInline> is not the same shape as <codeInline>array1</codeInline> ")
		_T("then cyclic indexing is used. ")
		_T("If <codeInline>outer</codeInline> is true then a two dimensional array with rows equal to ")
		_T("the number of elements of <codeInline>array1</codeInline> and columns equal to ")
		_T("the number of elements of <codeInline>array2</codeInline> is returned with the ")
		_T("result of <codeInline>function</codeInline> called for each pair in the cartesian product ")
		_T("of the arrays. ")
		_T("If <codeInline>array1</codeInline> is a handle, the underlying array is modified. ")
	)
);
xfp* WINAPI
xll_udf_apply2(double f, xfp* pa1, xfp* pa2, BOOL outer, BOOL trans)
{
#pragma XLLEXPORT
	xfp* px(0);

	try {
		px = xll_apply2(f, pa1, pa2, outer, trans);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return px;
}

#ifdef _DEBUG

static AddInX xai_test_apply0(
	_T("?xll_test_apply0"), XLL_DOUBLEX,
	_T("TEST.APPLY0"), _T("")
);
double WINAPI
xll_test_apply0(void)
{
#pragma XLLEXPORT
	static double count = 0;

	count = count + 1;

	return count;
}

static AddInX xai_test_apply1(
	_T("?xll_test_apply1"), XLL_DOUBLEX XLL_DOUBLEX,
	_T("TEST.APPLY1"), _T("Num")
);
double WINAPI
xll_test_apply1(double x)
{
#pragma XLLEXPORT

	return 2*x;
}

static AddInX xai_test_apply2(
	FunctionX(XLL_DOUBLEX, _T("?xll_test_apply2"), _T("TEST.APPLY2"))
	.Arg(XLL_DOUBLEX, _T("num1"), _T("is the first argument"))
	.Arg(XLL_DOUBLEX, _T("num2"), _T("is the second argument"))
);
double WINAPI
xll_test_apply2(double x, double y)
{
#pragma XLLEXPORT

	return x + y;
}

int
test_apply(void)
{
	try {

		OPERX o;

		o = ExcelX(xlfEvaluate, OPERX(_T("=UDF.APPLY0(TEST.APPLY0, 3)")));
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 3);
		ensure (o.columns() == 1);
		ensure (o[0] == 1);
		ensure (o[1] == 2);
		ensure (o[2] == 3);

		// constant function returning 1 + 2
		o = ExcelX(xlfEvaluate, OPERX(_T("=FUNCTIONAL.APPLY0(FUNCTIONAL.BIND(TEST.APPLY2, {1, 2}), 3)")));
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 3);
		ensure (o.columns() == 1);
		ensure (o[0] == 3);
		ensure (o[1] == 3);
		ensure (o[2] == 3);

		o = ExcelX(xlfEvaluate, OPERX(_T("=UDF.APPLY1(TEST.APPLY1, {1;2;3})")));
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 3);
		ensure (o.columns() == 1);
		ensure (o[0] == 2);
		ensure (o[1] == 4);
		ensure (o[2] == 6);

		// 2 + .
		o = ExcelX(xlfEvaluate, OPERX(_T("=FUNCTIONAL.APPLY1(FUNCTIONAL.BIND(TEST.APPLY2, 2, ), {1;2;3})")));
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 3);
		ensure (o.columns() == 1);
		ensure (o[0] == 3);
		ensure (o[1] == 4);
		ensure (o[2] == 5);

		o = ExcelX(xlfEvaluate, OPERX(_T("=UDF.APPLY2(TEST.APPLY2, {1,2,3}, {1,2})")));
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 1);
		ensure (o.columns() == 3);
		ensure (o[0] == 1 + 1);
		ensure (o[1] == 2 + 2);
		ensure (o[1] == 3 + 1);

		o = ExcelX(xlfEvaluate, OPERX(_T("=UDF.APPLY2(TEST.APPLY2, {1,2,3}, {1,2}, true)")));
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 3);
		ensure (o.columns() == 2);
		ensure (o(0, 0) == 1 + 1);
		ensure (o(2, 1) == 3 + 2);

		o = ExcelX(xlfEvaluate, OPERX(_T("=UDF.APPLY2(TEST.APPLY2, {1,2,3}, {1,2}, true, true)")));
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 2);
		ensure (o.columns() == 3);
		ensure (o(0, 0) == 1 + 1);
		ensure (o(1, 2) == 3 + 2);

		o = ExcelX(xlfEvaluate, OPERX(_T("=FUNCTIONAL.APPLY2(FUNCTIONAL.BIND(TEST.APPLY2, , ), {1,2,3}, {1,2})")));
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 1);
		ensure (o.columns() == 3);
		ensure (o[0] == 1 + 1);
		ensure (o[1] == 2 + 2);
		ensure (o[1] == 3 + 1);

		o = ExcelX(xlfEvaluate, OPERX(_T("=FUNCTIONAL.APPLY2(FUNCTIONAL.BIND(TEST.APPLY2, , ), {1,2,3}, {1,2}, true)")));
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 3);
		ensure (o.columns() == 2);
		ensure (o(0, 0) == 1 + 1);
		ensure (o(2, 1) == 3 + 2);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
//static Auto<OpenAfter> xao_apply(test_apply);

#endif // _DEBUG