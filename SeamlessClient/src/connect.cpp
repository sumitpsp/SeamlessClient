#include <iostream>
#include <cstring>
#include <arpa/inet.h>
#include <netdb.h>
#include <netinet/in.h>
#include <sys/types.h>
#include <ifaddrs.h>
#include "connect.h"
#include "cc.h"
#include "test_util.h"
#include "messages.h"


using namespace std;

#define MAXHOSTNAME 100

CConnect::CConnect() {
	server_ip = SERVER_IP;
	server_port = SERVER_PORT;
}

char* print_private_ip() {
    struct ifaddrs * ifAddrStruct=NULL;
    struct ifaddrs * ifa=NULL;
    void * tmpAddrPtr=NULL;

    getifaddrs(&ifAddrStruct);

	char ip[NI_MAXHOST];

    for (ifa = ifAddrStruct; ifa != NULL; ifa = ifa->ifa_next) {
        if (ifa ->ifa_addr->sa_family==AF_INET) { // check it is IP4
            // is a valid IP4 Address
            tmpAddrPtr=&((struct sockaddr_in *)ifa->ifa_addr)->sin_addr;
            char addressBuffer[INET_ADDRSTRLEN];
            inet_ntop(AF_INET, tmpAddrPtr, addressBuffer, INET_ADDRSTRLEN);
		
	
			if(!(strcmp(ifa->ifa_name,"en0"))) 
				strcpy(ip, addressBuffer);
			
			//For Cell Phones--Cell Network Address
			if(!(strcmp(ifa->ifa_name, "pdp_ip0"))) 
				strcpy(ip, addressBuffer);
		}		
	}

    if (ifAddrStruct!=NULL) freeifaddrs(ifAddrStruct);

	return ip;
}

int CConnect::connect_to_server() {
	int i;

	struct sockaddr_in my_addr;
	socklen_t sl =  sizeof(my_addr);
	char hostname[MAXHOSTNAME+1];
	struct sockaddr_in sa;
	struct hostent *hp;
	memset(&sa, 0, sizeof(struct sockaddr_in));
	gethostname(hostname, MAXHOSTNAME);           /* who are we? */
        hp= gethostbyname(hostname);  
 	if (hp == NULL)                             /* we don't exist !? */
    		return(-1);
  	sa.sin_family= hp->h_addrtype;              /* this is our host address */
  	sa.sin_port= htons(0);                /* this is our port number */
  	udt_client = UDT::socket(AF_INET, SOCK_STREAM, 0); /* socket to connect client to server */
	udt_server = UDT::socket(AF_INET, SOCK_STREAM, 0); /* socket to receive messages*/

  	if (UDT::bind(udt_client,(struct sockaddr *)&sa,sizeof(struct sockaddr_in)) == UDT::ERROR ) {
  		cout << "bind: " << UDT::getlasterror().getErrorMessage() << endl;
      		return 0;
	}
	
	
	char* ip = print_private_ip();
	int port;
	struct sockaddr_in sin;
    int len = sizeof(sin);
	if(UDT::getsockname(udt_client, (struct sockaddr *)&sin, &len) == 0) {
        port = ntohs(sin.sin_port);
        cout << "Private ip is " << ip << " : " <<port<< " " << strlen(ip) <<endl;
    }
    else {
        cout<< "Couldn't find port"<<endl;
        return 0;
    }
	
	sa.sin_port = htons(port);
	if (UDT::bind(udt_server,(struct sockaddr *)&sa,sizeof(struct sockaddr_in)) == UDT::ERROR ) {
        cout << "bind: " << UDT::getlasterror().getErrorMessage() << endl;
            return 0;
    }
	if(UDT::getsockname(udt_server, (struct sockaddr *)&sin, &len) == 0) {
        port = ntohs(sin.sin_port);
        cout << "Private ip is " << ip << " : " <<port<< " " << strlen(ip) <<endl;
    }
    else {
        cout<< "Couldn't find port"<<endl;
        return 0;
    }

	if(UDT::ERROR == UDT::listen(udt_server, 10)) {
        cout << "Listening error: "<<UDT::getlasterror().getErrorMessage() << endl;
    }
	
	
	handle_listener(udt_server);
	struct addrinfo hints, *local, *peer;
	
	memset(&hints, 0, sizeof(struct addrinfo));
	
	hints.ai_flags = AI_PASSIVE;
   	hints.ai_family = AF_INET;
   	hints.ai_socktype = SOCK_STREAM;
   	//hints.ai_socktype = SOCK_DGRAM;

	if (0 != getaddrinfo(server_ip, server_port, &hints, &peer))
   	{
      		cout << "incorrect server/peer address. " << server_ip << ":" << server_port << endl;
      		return 0;
   	}
	

	   // connect to the server, implict bind
   	if (UDT::ERROR == UDT::connect(udt_client, peer->ai_addr, peer->ai_addrlen))
   	{
   	   	cout << "connect: " << UDT::getlasterror().getErrorMessage() << endl;
      		return 0;
   	}

	cout<<"Connected to server"<<endl;
	


	char* buffer = NULL;

	len = reportip_payload(ip, port, NAME, &buffer);
	int temp;
    memcpy(&temp, buffer, sizeof(int));
    cout<<"Sent length is "<< temp<<endl;

	int ss = 0, ssize = 0;
	while (ssize < len) {
		if(UDT::ERROR == (ss = UDT::send(udt_client, buffer + ssize, len - ssize, 0)))
		{
			cout << "Send Error: "<< UDT::getlasterror().getErrorMessage() << endl;
			break;
		}
		ssize += ss;
	}
	//int temp;
	//memcpy(&temp, buffer, sizeof(int));
	//cout<<"Sent length is "<< temp<<endl;
	int a = temp;
	cout <<a;
	cout << "Sent size is " <<ssize<<endl;
	
	//free(buffer);
	UDT::close(udt_client);
}

int CConnect::handle_listener(UDTSOCKET socket) {
	int efd = UDT::epoll_create();
	if(efd == -1) {
		cerr<<"Epoll Create"<<endl;
	}
	UDT::epoll_add_usock(efd, socket);
	
	set<UDTSOCKET> readfds;

	while(true) {
		UDT::epoll_wait(efd, &readfds, NULL, -1);
	}
}

//int CConnect::reportip_payload_size(char 

int CConnect::reportip_payload(char* ip, int port, char* name, char** buffer) {
	
	int size_of_ip = strlen(ip) + 1;
	int size_of_name = strlen(name) + 1;
	
	int type = PEERINFO;

	//Size of Type + Size of IP Length + Size of IP String + Size of Port + Size of Name Length + Size of Name
	int payload_size = 	sizeof(type) + sizeof(size_of_ip) + size_of_ip + sizeof(port) + sizeof(size_of_name) + size_of_name;

	//Total Size = Length Header + Payload Size
	int total_size = sizeof(payload_size) + payload_size;
	
	cout<<"Payload Size is"<<total_size<<endl;;

	*buffer = (char*) (malloc(total_size));

	int added_size = 0;	
	char* message = *buffer;

	//Adding payload size
	memcpy(message, &payload_size, sizeof(payload_size));
	message+=sizeof(payload_size);
	added_size+=sizeof(payload_size);


	//Adding type of message
	memcpy(message, &type, sizeof(type));
	message+=sizeof(type);
	added_size+=sizeof(type);

	//Add size of ip length
	memcpy(message, &size_of_ip, sizeof(size_of_ip));
	message+=sizeof(size_of_ip);
	added_size+=sizeof(size_of_ip);

	//Add ip
	memcpy(message, ip, size_of_ip);
	message+=size_of_ip;
	added_size+=size_of_ip;

	//Add port
	memcpy(message, &port, sizeof(port));
	message+=sizeof(port);
	added_size+=sizeof(port);

	//Add size of Name
	memcpy(message, &size_of_name, sizeof(size_of_name));
	message+=sizeof(size_of_name);
	added_size+=sizeof(size_of_name);
	//Add name
	memcpy(message, name, sizeof(name));
	message+=size_of_name;
	added_size+=size_of_name;

	cout << "Added size is " << added_size <<endl;
	return total_size;

}

int main() {
	UDTUpDown _udt_;
	CConnect* connection = new CConnect();
	connection->connect_to_server();
	return 0;
}




