# USE CASE
When there are different kind of model outputs, for example, Aermod for construction dust, Caline for vehicular emissions, and Path for background contribution, we want to summarize the results to get the aggregated impacts.

# BASIC IDEA
#Not apply to path because it is grid specific.

The basic idea is to sum the hourly emission in different model together, then reuse the program in project aermod to summarize the results according to legislative requirements.

For the breakdown, the program will run each kind of model outputs once.