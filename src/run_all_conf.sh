#!/bin/bash

for i in "0" "0.1" "0.2" "0.3" "0.4" "0.5" "0.6" "0.7" "0.8" "0.9" "1"; do
	echo "Generating request data for target conflict rate ${i}"
	python3 generate_requests.py ../data/data_generated_0.8 ../data/data_generated_conf_${i} --on_probability 0.8 --off_probability 0 --conflict_probability ${i}
	cp ../data/data_generated_0.8/*.cmpl ../data/data_generated_conf_${i}
	echo "Solving all strategies for target conflict rate ${i}"
	python3 solve_all.py ../data/data_generated_conf_${i} ../results/output_generated_conf_${i}
	python3 eval_solver_runs.py ../data/data_generated_conf_${i} ../results/output_generated_conf_${i} ../results/results_generated_conf_${i}.xlsx
done
