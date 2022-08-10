PYTHON = python3
SETUNITTEST = -m unittest -v -b

# E501 - Line too long
# TODO remove R0201, R0801
linter : flake8_linter pylint_linter pylint_linter_tests

flake8_linter:
	flake8 tests sheets --ignore E501,W503

pylint_linter:
	python3 -m pylint sheets/ --good-names=op,ii,jj,cc,vv --disable=C0301,R0201,R0801

pylint_linter_tests:
	python3 -m pylint tests/ --good-names=wb,ii,jj,cc,vv --disable=C0301,R0801,C0103,R0902

all_tests :
	${PYTHON} ${SETUNITTEST} tests/test*

stress_tests :
	${PYTHON} ${SETUNITTEST} tests/stresstest*

curr_test:
	${PYTHON} ${SETUNITTEST} tests/testfunctions.py

find_todos:
	grep -rwn "TODO" docs sheets tests > TODOs.txt
