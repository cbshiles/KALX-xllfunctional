// functional.cpp - STL <functional>
#include "xll/utility/file.h"
#include "functional.h"

using namespace xll;

#ifdef _DEBUG
static AddInX xai_array_doc(
	DocumentX(CATEGORY)
	.Documentation(
		XML_FILE("functional.xml")
	)
);	
#endif

static AddInX xai_nan(
	FunctionX(XLL_FPX, _T("?xll_nan"), _T("NAN"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns an IEEE 64-bit Not A Number value."))
	.Documentation()
);
xfp* WINAPI
xll_nan(void)
{
#pragma XLLEXPORT
	static xfp x = {1, 1};
	x.array[0] = std::numeric_limits<double>::quiet_NaN();

	return &x; // direct NaN gets converted to #NUM!
}
static AddInX xai_isnan(
	FunctionX(XLL_BOOLX, _T("?xll_isnan"), _T("ISNAN"))
	.Arg(XLL_DOUBLEX, _T("Number"), IS_NUM)
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns TRUE if Number is an IEEE 64-bit Not A Number value."))
	.Documentation()
);
BOOL WINAPI
xll_isnan(double x)
{
#pragma XLLEXPORT

	return _isnan(x);
}

static AddInX xai_sum(
	FunctionX(XLL_DOUBLEX, _T("?xll_sum"), _T("FUNCTION.ADD"))
	.Arg(XLL_DOUBLEX, _T("arg1"), _T("is the first argument"))
	.Arg(XLL_DOUBLEX, _T("arg2"), _T("is the second argument. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns arg1 + arg2 "))
	.Documentation()
);
double WINAPI
xll_sum(double x1, double x2)
{
#pragma XLLEXPORT

	return x1 + x2;
}

static AddInX xai_difference(
	FunctionX(XLL_DOUBLEX, _T("?xll_difference"), _T("FUNCTION.SUB"))
	.Arg(XLL_DOUBLEX, _T("arg1"), _T("is the first argument"))
	.Arg(XLL_DOUBLEX, _T("arg2"), _T("is the second argument. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns arg1 - arg2 "))
	.Documentation()
);
double WINAPI
xll_difference(double x1, double x2)
{
#pragma XLLEXPORT

	return x1 - x2;
}

static AddInX xai_product(
	FunctionX(XLL_DOUBLEX, _T("?xll_product"), _T("FUNCTION.MUL"))
	.Arg(XLL_DOUBLEX, _T("arg1"), _T("is the first argument"))
	.Arg(XLL_DOUBLEX, _T("arg2"), _T("is the second argument. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns arg1 * arg2 "))
	.Documentation()
);
double WINAPI
xll_product(double x1, double x2)
{
#pragma XLLEXPORT

	return x1 * x2;
}

static AddInX xai_quotient(
	FunctionX(XLL_DOUBLEX, _T("?xll_quotient"), _T("FUNCTION.DIV"))
	.Arg(XLL_DOUBLEX, _T("arg1"), _T("is the first argument"))
	.Arg(XLL_DOUBLEX, _T("arg2"), _T("is the second argument. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns arg1 / arg2 "))
	.Documentation()
);
double WINAPI
xll_quotient(double x1, double x2)
{
#pragma XLLEXPORT

	double x = x1/x2;

	return x;
}
static AddInX xai_pow(
	FunctionX(XLL_DOUBLEX, _T("?xll_pow"), _T("FUNCTION.POW"))
	.Arg(XLL_DOUBLEX, _T("arg1"), _T("is the first argument"))
	.Arg(XLL_DOUBLEX, _T("arg2"), _T("is the second argument. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns arg1 / arg2 "))
	.Documentation()
);
double WINAPI
xll_pow(double x1, double x2)
{
#pragma XLLEXPORT

	double x = pow(x1, x2);

	return x;
}

static AddInX xai_mod(
	FunctionX(XLL_LPOPERX, _T("?xll_mod"), _T("FUNCTION.MOD"))
	.Arg(XLL_LONGX, _T("arg1"), _T("is the first argument"))
	.Arg(XLL_LONGX, _T("arg2"), _T("is the second argument. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns arg1 % arg2 "))
	.Documentation(
		_T("This function is not the same as the Excel built-in function <codeInline>POW</codeInline>. ")
	)
);
LPOPERX WINAPI
xll_mod(LONG x1, LONG x2)
{
#pragma XLLEXPORT
	static OPERX o;

	if (x2 != 0) {
		o.xltype = xltypeNum;
		o.val.num = x1 % x2;
	}
	else {
		o = ErrX(xlerrValue);
	}

	return &o;
}

static AddInX xai_max(
	FunctionX(XLL_DOUBLEX, _T("?xll_max"), _T("FUNCTION.MAX"))
	.Arg(XLL_DOUBLEX, _T("arg1"), _T("is the first argument"))
	.Arg(XLL_DOUBLEX, _T("arg2"), _T("is the second argument. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns the maximum of arg1 and arg2. "))
	.Documentation()
);
double WINAPI
xll_max(double x1, double x2)
{
#pragma XLLEXPORT

	return max(x1,x2);
}

static AddInX xai_min(
	FunctionX(XLL_DOUBLEX, _T("?xll_min"), _T("FUNCTION.MIN"))
	.Arg(XLL_DOUBLEX, _T("arg1"), _T("is the first argument"))
	.Arg(XLL_DOUBLEX, _T("arg2"), _T("is the second argument. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns the minimum of arg1 and arg2. "))
	.Documentation()
);
double WINAPI
xll_min(double x1, double x2)
{
#pragma XLLEXPORT

	return min(x1,x2);
}

static AddInX xai_not(
	FunctionX(XLL_WORDX, _T("?xll_not"), _T("FUNCTION.NOT"))
	.Arg(XLL_WORDX, _T("Arg"), _T("is an integer argument. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns the bitwise logical compliment of Arg."))
	.Documentation(_T("This function is the same as the C ~ operator on integers."))
);
WORD WINAPI
xll_not(WORD x)
{
#pragma XLLEXPORT

	return ~x;
}

static AddInX xai_and(
	FunctionX(XLL_WORDX, _T("?xll_and"), _T("FUNCTION.AND"))
	.Arg(XLL_WORDX, _T("arg1"), _T("is the first argument"))
	.Arg(XLL_WORDX, _T("arg2"), _T("is the second argument. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns the bitwise logical and of arg1 and arg2."))
	.Documentation(_T("This is the same as the C &#38; operator on integers."))
);
WORD WINAPI
xll_and(WORD x1, WORD x2)
{
#pragma XLLEXPORT

	return x1 & x2;
}

static AddInX xai_or(
	FunctionX(XLL_WORDX, _T("?xll_or"), _T("FUNCTION.OR"))
	.Arg(XLL_WORDX, _T("arg1"), _T("is the first argument"))
	.Arg(XLL_WORDX, _T("arg2"), _T("is the second argument. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns the bitwise logical or of arg1 and arg2."))
	.Documentation(_T("This is the same as the C | operator on integers."))
);
WORD WINAPI
xll_or(WORD x1, WORD x2)
{
#pragma XLLEXPORT

	return x1 | x2;
}

static AddInX xai_xor(
	FunctionX(XLL_WORDX, _T("?xll_xor"), _T("FUNCTION.XOR"))
	.Arg(XLL_WORDX, _T("arg1"), _T("is the first argument"))
	.Arg(XLL_WORDX, _T("arg2"), _T("is the second argument. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns the bitwise exclusive or of arg1 and arg2."))
	.Documentation(_T("This is the same as the C ^ operator on integers."))
);
WORD WINAPI
xll_xor(WORD x1, WORD x2)
{
#pragma XLLEXPORT

	return x1 ^ x2;
}

static AddInX xai_shift(
	FunctionX(XLL_WORDX, _T("?xll_shift"), _T("FUNCTION.SHIFT"))
	.Arg(XLL_WORDX, _T("Arg"), _T("is the argument whose bits are to be shifted."))
	.Arg(XLL_SHORTX, _T("N"), _T("is number of bits to shifted to the left if positive or right if negative. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns Arg &#60;&#60; N or Arg &#62;&#62; N. "))
	.Documentation()
);
WORD WINAPI
xll_shift(WORD x, SHORT n)
{
#pragma XLLEXPORT

	return n > 0 ? x << n : n < 0 ? x >> -n : x;
}

static AddInX xai_equals(
	FunctionX(XLL_BOOLX, _T("?xll_equals"), _T("FUNCTION.EQUALS"))
	.Arg(XLL_LPOPERX, _T("arg1"), _T("is the first argument"))
	.Arg(XLL_LPOPERX, _T("arg2"), _T("is the second argument. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns true if arg1 is equal to arg2. "))
	.Documentation()
);
BOOL WINAPI
xll_equals(LPOPERX px1, LPOPERX px2)
{
#pragma XLLEXPORT

	return *px1 == *px2;
}

static AddInX xai_identity(
	FunctionX(XLL_BOOLX, _T("?xll_identity"), _T("FUNCTION.IDENTITY"))
	.Arg(XLL_LPOPERX, _T("Arg"), _T("is the argument to be returned. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns Arg."))
	.Documentation()
);
LPOPERX WINAPI
xll_identity(LPOPERX px)
{
#pragma XLLEXPORT

	return px;
}
