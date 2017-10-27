// functional.h - user defined functiona
// Copyright (c) 2006-2010 KALX, LLC. All rights reserved. No warranty is made.
#pragma once
//#define EXCEL12
#include "xll/xll.h"

#ifndef CATEGORY
#define CATEGORY _T("Functional")
#endif

#define IS_ARRAY_OR_HANDLE _T("is an array of numbers or a handle to an array")
#define IS_ARRAY _T("is an array of numbers")
#define IS_UDF _T("is the register id of a user defined function")
#define IS_FUNCTIONAL _T("is a handle returned by FUNCTIONAL.BIND")
#define IS_ROWS _T("is the number of rows in the array")
#define IS_COLUMNS _T("is the number of columns in the array")
#define IS_NUM _T("is a floating point number")

typedef xll::traits<XLOPERX>::xcstr xcstr;
typedef xll::traits<XLOPERX>::xfp xfp;
typedef xll::traits<XLOPERX>::xword xword;

namespace function {

	// unwrap OPER handles if necessary
	inline OPERX 
	get(const OPERX& x)
	{
		if (x.xltype == xltypeNum) {
			xll::handle<OPERX> x_(x.val.num);
			if (x_)
				return *x_;
		}

		return x; // might be an array of handles
	}

	class bind {
		OPERX fx_; // f, x?!
	public:

		// f is the register id of a UDF
		bind(HANDLEX f, const OPERX& x)
			: fx_(1 + x.size(), 1)
		{
			// set up for call to Excelv
			fx_[0] = f;
			for (xword i = 0; i < x.size(); ++i) {
				fx_[i + 1] = x[i];
			}
		}
		OPERX call(void) const
		{
			return xll::Excelv<XLOPERX>(xlUDF, fx_.size(), fx_.begin());
		}
		// call with missing args
		OPERX call(const OPERX& _x) const
		{
			OPERX fx = fx_;

			if (_x.size() > 0) {
				xword j = 0;
				for (xword i = 1; i < fx.size(); ++i) {
					if (fx[i].xltype != xltypeNum) {
						fx[i] = get(_x[j++]);
					}
					else if (fx[i].xltype == xltypeNum) {
						xll::handle<bind> h(fx[i].val.num);
						if (h)
							fx[i] = h->call(get(_x[j++])); //recursive
					}
				}
				ensure (j == _x.size());
			}

			return xll::Excelv<XLOPERX>(xlUDF, fx.size(), fx.begin());
		}
	};
	/*

		// bound function + args
		struct args : public OPERX {
			args()
				: OPERX()
			{ }
			args(xword n)
				: OPERX(n, 1)
			{ }
			args& bind(const OPERX& x)
			{
				for (xword i = 1; i < size(); ++i) {
					if (operator[](i).xltype == xltypeMissing) {
						operator[](i) = x;

						return *this;
					}
				}
			
				throw std::runtime_error("func::bind: no arguments to bind");
			}
		};

		class func {
			xword n_; // arity
			xword m_; // number of missing args
			args  x_; // function and possibly missing args
		public:
			func(double f, const OPERX& x)
				: m_(0)
			{
				const xll::ArgsX* pa;
				ensure (0 != (pa = xll::ArgsMapX::Find(f)));
				ensure ((n_ = pa->ArgCount()) <= 4);
				ensure (n_ <= x.size()); // x can be too big, extra args ignored

				x_.resize(n_ + 1, 1);
				x_[0] = f;
				for (xword i = 1; i <= n_; ++i) {
					x_[i] = x[i - 1];
					if (x_[i].xltype == xltypeMissing)
						++m_;
				}
			}
			xword arity(void) const
			{
				return n_;
			}
			xword missing(void) const
			{
				return m_;
			}
			// return by value
			args args(void) const
			{
				return x_;
			}
			static OPERX udf(const OPERX& x)
			{
				return xll::Excelv<XLOPERX>(xlUDF, x.size(), &x[0]);
			}
			// make a copy of args
			OPERX eval(OPERX x
				, const OPERX& x1 = MissingX()
				, const OPERX& x2 = MissingX()
				, const OPERX& x3 = MissingX()
				, const OPERX& x4 = MissingX()
				)
			{
				// x is really args
				ensure (x.size() == 1 + arity());

				for (xword i = 1; i <= arity(); ++i) {
					xword n = x.size() - i;
					if (x[n].xltype == xltypeNum) {
						xll::handle<func> f(x[n].val.num);
						if (f) {
							// call using nearest args
							xword m = f->missing();
							if (m == 0)
								x[n] = udf(f->args());
							else if (m == 1)
								x[n] == udf(f->args().bind(x1));
							else if (m == 2)
								x[n] == udf(f->args().bind(x1).bind(x2));
							else if (m == 3)
								x[n] == udf(f->args().bind(x1).bind(x2).bind(x3));
							else if (m == 4)
								x[n] == udf(f->args().bind(x1).bind(x2).bind(x3).bind(x4));
							else
								throw std::runtime_error("func::eval: subargument arity is too large");
						}
					}
				}

				return udf(x);
			}
			OPERX call()
			{
				return eval(args());
			}
			OPERX call(const OPERX& x1)
			{
				return eval(args(), x1);
			}
			OPERX call(const OPERX& x1, const OPERX& x2)
			{
				return eval(args(), x1, x2);
			}
			OPERX call(const OPERX& x1, const OPERX& x2, const OPERX& x3)
			{
				return eval(args(), x1, x2, x3);
			}
			OPERX call(const OPERX& x1, const OPERX& x2, const OPERX& x3, const OPERX& x4)
			{
				return eval(args(), x1, x2, x3, x4);
			}
		};
	*/
} // function