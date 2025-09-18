# Research into methods

Owner: Edward Brady

## [Constraint satisfaction problems:](Experimenting%20with%20CSP%2020b713b1f81d800fae3be4bdc71357b1.md)

[Description of CSP and how they solved here](https://medium.com/@kanchanakanta/constraint-satisfaction-problems-csp-766f3ddeed3f)

These are often used for the scheduling, they allow for hard and soft constraints and to put preferences on each constraint.

Pros:

- Can do hard and soft constraints
- Can do complex constraints

Cons:

- May not be great at scaling
- May not be quick

Useful packages:

[https://github.com/python-constraint/python-constraint](https://github.com/python-constraint/python-constraint)

[OR-Tools  |  Google for Developers](https://developers.google.com/optimization)

[https://github.com/budrus123/GraduationProject/blob/master/project-overview/GradProjectFinal2018.pdf](https://github.com/budrus123/GraduationProject/blob/master/project-overview/GradProjectFinal2018.pdf) - a githup repo for very similar project written in java 

## Integer Linear Programming (ILP):

Pros:

- Can be used to optimize the scheduling
- Works well for simple clear objectives

Cons:

- Harder to express non linear or conditional constraints
- Can’t do soft constraints

Useful packages:

[https://pypi.org/project/PuLP/](https://pypi.org/project/PuLP/)

## Heuristics and Metaheuristics:

Pros:

- Very fast and scalable
- Can include soft constraints

Cons:

- More difficult to write and tune
- Won’t necessarily enforce the hard constraints

Useful packages:

[DEAP](https://github.com/DEAP/deap)

[https://pymoo.org/](https://pymoo.org/)

## Key take aways:

1. Start with CSP looking into packages
2. Optimize with ILP if needed