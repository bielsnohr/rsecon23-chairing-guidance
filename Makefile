SESSIONS_FILE = sessions.csv
TALKS_FILE = talks.csv
OUT_DIR = infosheets

.PHONY : clean infosheets

infosheets : ${SESSIONS_FILE} ${TALKS_FILE}
	mkdir -p ${OUT_DIR}
	python chairing.py $^ ${OUT_DIR}

clean :
	rm -r ${OUT_DIR}
