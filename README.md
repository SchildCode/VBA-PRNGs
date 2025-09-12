# VBA-PRNGs
Pseudoranom number generators ported to VBA, and code for validating with TestU01 test battery
I have used pseudorandom and quasirandom number generators for over 30 years, for Monte Carlo simulations, optimization, and numerical quadrature.
I have translated a number of PRNGs to VBA. Recently I re-evaluated which PRNGs I should use in VBA, and the results are summarized in the table below in order of speed:

| PRNG name | Time to generate 1&nbsp;million&nbsp;numbers | SmallCrush test | Source |
| --------- | -------------------------- | --------------- | ------ |
| rnd()     | 0.07 s                     | <span style="color: red;">fail</span> | native Excel VBA function |
| Bryc32    | 0.08 s                     | pass             | "Bryc" variant of SFC32 |
| fminx32   | 0.09 s                     | pass             | Canonical SplitMix32 with Murmur3 fmix32 xmxmx-mixer/finisher |
| mix32     | 0.09 s                     | pass             | SplitMix32 with xmxmx-mixer optimized by Hash-Prospector |
| SFC32     | 0.09 s                     | pass             | Small Fast Counting (SFC) by Chris Doty-Humphrey |
| LFIB4     | 0.22 s                     | pass             | 4-lagged Fibonacci generator (Marsaglia) |
| RANDBETWEEN() | 0.23 s                 | pass             | Array versjon of Excel worksheet function RAND | 
| PCG32     | 0.51 s                     | pass             | PCG32 (XSH-RR),  Permuted Congruential Generator | 
| MTran     | 0.81 s                     | pass             | Mersenne Twister |
| RAND()    | 41.40 s                    | pass             | =Evaluate("=RAND()") in VBA |

## Conclusion

* Use **SFC32** for most PRNG tasks (all except studies to be published with peer-review), in partocular those requiring speed. The source code is in file [SFC32_module.bas](https://github.com/SchildCode/VBA-PRNGs/blob/main/SFC32_module.bas).
* Use **PCG32** for those PRNG tasks requiring maximum statistical quality, those to be published scientifically. The source code is in file [PCG32_module.bas](https://github.com/SchildCode/VBA-PRNGs/blob/main/PCG32_module.bas).

Code of the remaining PRNGs is in file [otherPRNGs_module.bas](https://github.com/SchildCode/VBA-PRNGs/blob/main/otherPRNGs_module.bas).<br>
My code for testing all the PRNGs on both ENT and the TestU01 test batteries (SmallCrush, Crush & BigCrush) are included in file tests_module.bas.


  
