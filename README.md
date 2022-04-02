# Actuarial_Projected_Claims
Calculate the projected claims for the full duration of the policy for any given policyholder using Excel formulas and VBA.

#### Background
An insurance provider offers a healthcare policy that covers medical, dental, and vision expenses.

This healthcare policy provides coverage for a maximum of 25 years or until the policyholder reaches age 70. This insurance provider does not offer this policy to individuals under 20 years old.

The average claim amount for a male is $687 and the average claim amount for a female is $492.

The probability of making a claim by age bracket for males is as follows:

      20s	    30s	    40s	    50s	    60s
Male	1.00%	  1.50%	  3.00%	  5.00%	  8.00%

Females are 25% less likely to make a claim at any given age. So, females will always have a lower probability of making a claim.

The probability of making a claim is 50% more for those that are considered unhealthy.

#### Task
My task is to assist the insurance provider by calculating the projected claims for the full duration of the policy for any given policyholder.

The tool should be able to accept inputs about the policyholder (gender [M/F], age of the policyholder [20 to 69], and health condition [healthy/unhealthy])

The output of the tool should be a projection of the expected claim amounts for the policyholder in Year 1, Year 2, Year 3, … etc. of the policy.

Note: Assume that a policyholder always purchases the policy on their birthday, so the policyholder will be the same age for the entire duration of each year.  For example, if a 25 year old purchases a policy on their 25th birthday, their age for Year 1 of the policy is 25, and their age for Year 2 of the policy is 26, etc..

#### Result
The input cells are programed only to accept certain values: (gender [M/F], age of the policyholder [20 to 69], and health condition [healthy/unhealthy])

Based on the age of the policy holder, formulas populate the age category for each year of the policy.

The user then clicks the "Calculate" button and the VBA macro I wrote calculates the projected claim amount to be paid each year of the policy and populates those values in the Projected Claims row.

The total projected amount paid for claims over the life of the policy is displayed below the "Calculate" button.

