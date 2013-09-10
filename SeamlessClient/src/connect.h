#include <udt.h>

#define SERVER_IP "localhost" //"54.227.241.12"

#define SERVER_PORT "11111"
#define NAME "sumitpc"

class CConnect {
        char* server_ip;
		char* server_port;
		int local_socket;
		UDTSOCKET udt_client, udt_server;
		int reportip_payload(char* ip, int port, char* name, char** buffer);
		int getpeer_payload();
		static int handle_listener(UDTSOCKET socket);
	public:
        CConnect ();
        int refresh_server();
        int connect_to_server();
        int connect_to_peer();
        int send_to_peer();
        int receive_from_peer();
};

