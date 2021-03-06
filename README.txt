MacroStats
A collection of statistics-related functions for Excel's VBA macro language.

By Iain Dunning, 2011
http://www.iaindunning.com
https://github.com/IainNZ/MacroStats

------------------------------------------------------------------------------
NEWS
------------------------------------------------------------------------------

* 18th May 2011
	- Project started, added some functions from a simulation project, tidied
      them up and added some tests.
	- Lots more to come!

------------------------------------------------------------------------------
PROJECT LAYOUT
------------------------------------------------------------------------------

* MacroStats.bas
	- The library itself, IMPORT THIS INTO YOUR PROJECTS!
* MacroStats.xlsm
	- The workbook used to develop the libary.
	- Don't need to add this to you to your project.
* MacroStatsTest.bas
	- Mainly used to aid development, not necessary to include in your
          projects.
	- Used to aid development by ensuring that the functions work as
          promised. Ff you encounter a problem and report it to me, I'll
          add a test that should ensure the problem is never reintroduced.

------------------------------------------------------------------------------
LIBRARY CONTENTS
------------------------------------------------------------------------------
See MacroStats.bas for full descriptions of functions, as well as samples
that show how they can be used.

Probability-related
===================
* SampleDiscreteCDF
	Inputs: CDF()
	Generate a random integer from a cumulative distribution function.

* SampleDiscreteCDFon2D
	Inputs: CDF(), FirstIndex
	Generate a random integer from a cumulative distribution function.
	Use this when you have a 2D array, where the CDF is in the 2nd dimension.

* SampleDiscreteCDFon3D
	Inputs: CDF(), FirstIndex, SecondIndex
	Generate a random integer from a cumulative distribution function.
	Use this when you have a 3D array, where the CDF is in the 3rd dimension.
	
* FlipCoin
	Inputs: Probability
	Returns a true with the provided probability.
	
Distribution-related
====================
* FitNormalDistributionToData
	Inputs: data()
    Outputs: mean, stddev
	Fits a normal distribution to a data set provided by the user.
	Uses Maximum Likelihood Estimation.

* FitNormalDistributionToPercentiles
	Inputs: X1, P1, X2, P2
	Outputs: mean, stddev
	Fits a normal distribution to two percentiles.
	e.g.
	50th percentile is 60 => X1 = 60, P1 = 0.5
	84th percentile is 80 >= X2 = 80, P2 = 0.84
	-> mean = 60, stddev = 20
	
* RandomFromNormal
	Inputs: mean, stddev
	Uses Box-Mueller to generate a random normal number.
	
* FitGammaDistributionToData
	Inputs: data(),
    Outputs: shapeP, scaleP
	Fits a gamma distribution to a data set provided by the user.
	Uses Maximum Likelihood Estimation with Newton-Rhapson.

* FitGammaDistributionToPercentiles
	Inputs: X1, P1, X2, P2
	Outputs: shapeP, scaleP
	e.g.
	50th percentile is 5000 => X1 = 5000, P1 = 0.5
	84th percentile is 6500 >= X2 = 6500, P2 = 0.84
	-> shape = 13.5, scale = 378.9