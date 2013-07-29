#define REPORTIP 1
#define PEERINFO 2

typedef struct  {
    int length;
    int type;
} message;

typedef struct {
	char* ip;
	int length;
} message_reportip;

