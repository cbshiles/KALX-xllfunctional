// oper.cpp - create an oper to pass to bind
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "functional.h"

using namespace xll;

typedef traits<XLOPERX>::xfp xfp;

static AddInX xai_nil(
	FunctionX(XLL_LPOPERX, _T("?xll_nil"), _T("NIL"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return an OPER of type xltypeNil."))
	.Documentation()
);
LPOPERX WINAPI
xll_nil()
{
#pragma XLLEXPORT
	static OPERX o = NilX();

	return &o;
}