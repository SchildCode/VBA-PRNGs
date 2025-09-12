# VBA-PRNGs
Pseudoranom number generators ported to VBA, and code for validating with TestU01 test battery
I have used pseudorandom and quasirandom number generators for over 30 years, for Monte Carlo simulations, optimization, and numerical quadrature.
I have translated a number of PRNGs to VBA. Recently I re-evaluated which PRNGs I should use in VBA, and the results are summarized in the table below in order of speed:

| PRNG name | Time to generate 1 million | SmallCrush test |
| --------- | -------------------------- | --------------- |
| rnd()     | 0.07 s                     | fail             |
| Bryc32    | 0.08 s                     | pass             |
| fminx32   | 0.09 s                     | pass             | 
| mix32     | 0.09 s                     | pass             |
| SFC32     | 0.09 s                     | pass             | 
| LFIB4     | 0.22 s                     | pass             | 
| RANDBETWEEN() | 0.23 s                 | pass             |
| PCG32     | 0.51 s                     | pass             |
| MTran     | 0.81 s                     | pass             |
| RAND()    | 41.40 s                    | pass             |

## Conclusion

* Use SFC32 for most PRNG tasks (all except studies to be published with peer-review)
* Use PCG32 for the remaining 5% of PRNG tasks, those to be published scientifically

  
