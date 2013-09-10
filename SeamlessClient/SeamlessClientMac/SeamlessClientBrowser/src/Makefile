CXX=g++
PLUGIN=seamlessPlugin.o
TARGETS=seamlessPlugin

%.o: %.cpp
	$(CXX) -I../npapi-sdk-read-only/headers -c $<

all: $(TARGETS)

seamlessPlugin: $(PLUGIN)
	$(CXX) $^ -o $@

.PHONY: clean

clean:
	rm -fr *.o 


