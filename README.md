Delphi Expression Evaluater
===========================

This library is based on the prExpr.pas unit:
- the original works of Martin Lafferty:
- Copyright:   1997 Production Robots Engineering Ltd, all rights reserved.
- Version:     1.03 31/1/98

Martin was good enough to grant non-exclusive usage and distribution rights to anyone in need of a great expression evaluator.
(see copyright notice inside prExpr.pas)

In 1997 prExpr was among the fastest and most reliable of them all and maybe still is.
As time went on I made some changes to the orriginal code and extended the base expression parser with two extra units!

The most important change I added to this library was the use of 'AsVariant' and the interfaced 'ParameterList'.
I also added the 'IValue' interface to the original TExpression object (because I only had a copy a that time), but later I noticed Martin also added this to the library with version 1.04.

Currently I still use this library extensively in production environments and have never had to need anything else for parsing expressions in Delphi applications.

I could not find any new work by Martin on the prExpr code, so I thought it would be nice to share his work with mine to the world on Git!.

I'll try to update the Wiki in time. For now it only contains some basics about the library and a reference to the documents in the [/Docs](https://github.com/ronaldhoek/ExprEval/tree/master/Docs) folder of the git project.

Please enjoy using the library.
