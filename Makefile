SESSIONS_FILE = sessions.csv
TALKS_FILE = talks.csv
OUT_DIR = infosheets

.PHONY : clean infosheets

infosheets : ${SESSIONS_FILE} ${TALKS_FILE}
	mkdir -p ${OUT_DIR}
	python3 chairing.py $^ ${OUT_DIR}

test: sessions_test.csv talks_test.csv
	mkdir -p test_output
	python3 chairing.py $^ test_output


clean :
	rm -r ${OUT_DIR}
