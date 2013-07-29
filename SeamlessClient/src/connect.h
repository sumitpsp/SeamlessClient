#include <udt.h>

#define SERVER_IP "54.227.241.12"
#define SERVER_PORT "11112"

class CConnect {
        char* server_ip;
		char* server_port;
		int local_socket;
		UDTSOCKET udt_client;
	public:
        CConnect ();
        int refresh_server();
        int connect_to_server();
        int connect_to_peer();
        int send_to_peer();
        int receive_from_peer();
};

