# VBA-PRNGs
Pseudoranom number generators ported to VBA, and code for validating with TestU01 test battery
I have used pseudorandom and quasirandom number generators for over 30 years, for Monte Carlo simulations, optimization, and numerical quadrature.
I have translated a number of PRNGs to VBA. Recently I re-evaluated which PRNGs I should use in VBA, and the results are summarized in the table below in order of speed:

| PRNG name | Time to generate 1&nbsp;million&nbsp;numbers | SmallCrush | Crush | BigCrush | Period | Source |
| --------- | -------------------------------------------- | ---------- | ----- | -------- | ------ | -------| 
| rnd()     | 0.068 s                     | ![Failed](https://img.shields.io/badge/Fail-red) | ![Failed](https://img.shields.io/badge/Fail-red) | ![Failed](https://img.shields.io/badge/Fail-red) | 2<sup>24</sup> ?| native Excel VBA function |
| SFC32     | 0.088 s                     | ![Passed](https://img.shields.io/badge/Pass-green) | ![Passed](https://img.shields.io/badge/Pass-green) | ![Passed](https://img.shields.io/badge/Pass-green) | ~2<sup>127</sup> | Small Fast Counting (SFC) by Chris Doty-Humphrey |
| fminx32   | 0.091 s                     | ![Passed](https://img.shields.io/badge/Pass-green) | ![Failed](https://img.shields.io/badge/Fail-red) | ![Failed](https://img.shields.io/badge/Fail-red) | 2<sup>32</sup> | Canonical SplitMix32 with Murmur3 fmix32 xmxmx-mixer/finisher |
| mix32     | 0.094 s                     | ![Passed](https://img.shields.io/badge/Pass-green) | ![Failed](https://img.shields.io/badge/Fail-red) | ![Failed](https://img.shields.io/badge/Fail-red) | 2<sup>32</sup> | SplitMix32 with xmxmx-mixer optimized by Hash-Prospector |
| RANDBETWEEN() | 0.209 s                 | ![Passed](https://img.shields.io/badge/Pass-green) | ![Passed](https://img.shields.io/badge/Pass-green) | ? | ? | Faster array-version of Excel worksheet function RAND() | 
| LFIB4     | 0.212 s                     | ![Passed](https://img.shields.io/badge/Pass-green) | ![Passed](https://img.shields.io/badge/Pass-green) | ![Failed](https://img.shields.io/badge/Fail-red) | ~2<sup>287</sup> | 4-lagged Fibonacci generator (Marsaglia) |
| PCG32     | 0.455 s                     | ![Passed](https://img.shields.io/badge/Pass-green) | ![Passed](https://img.shields.io/badge/Pass-green) | ![Passed](https://img.shields.io/badge/Pass-green) | 2<sup>64</sup> | PCG32 (XSH-RR),  Permuted Congruential Generator | 
| MTran     | 0.813 s                     | ![Passed](https://img.shields.io/badge/Pass-green) | ![Passed](https://img.shields.io/badge/Pass-green) | ![Failed](https://img.shields.io/badge/Fail-red) | 2<sup>19937</sup>-1 | Mersenne Twister, a.k.a. MT19937 |
| RAND()    | 41.40 s                    | ![Passed](https://img.shields.io/badge/Pass-green) | ![Passed](https://img.shields.io/badge/Pass-green) | ? | ? | =Evaluate("=RAND()") in VBA |

## Conclusion

* Use **SFC32** for most PRNG tasks (all except studies to be published with peer-review). It is very fast, has a long period (~2^128), and passes all tests in BugCrush. I have also coded an 'extra-small' version, **xSFC32**, with self-contained initialization and hard-coded optimixed seeding. The source code for both is in file [SFC32_module.bas](https://github.com/SchildCode/VBA-PRNGs/blob/main/SFC32_module.bas).
* Use **PCG32** for those PRNG tasks requiring maximum statistical quality, those to be published scientifically. It is slower than SFC32 but has a solid mathematical design rationale (LCG base + permutation step) and extensive empirical evidence of consistently excellent quality (TestU01, PractRand) and reproducibility. The source code is in file [PCG32_module.bas](https://github.com/SchildCode/VBA-PRNGs/blob/main/PCG32_module.bas).

Code of the remaining PRNGs is in file [otherPRNGs_module.bas](https://github.com/SchildCode/VBA-PRNGs/blob/main/otherPRNGs_module.bas).<br>
The above test results are from the specific VBA code in this GitHub repositry.  My code for testing all the PRNGs on both [ENT](https://cacert.at/random/) and the [TestU01](https://en.wikipedia.org/wiki/TestU01) test batteries (SmallCrush, Crush & BigCrush) are included in file [tests_module.bas](https://github.com/SchildCode/VBA-PRNGs/blob/main/tests_module.bas).


  
