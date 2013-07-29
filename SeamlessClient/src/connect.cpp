#include <iostream>
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

void print_private_ip() {
    struct ifaddrs * ifAddrStruct=NULL;
    struct ifaddrs * ifa=NULL;
    void * tmpAddrPtr=NULL;

    getifaddrs(&ifAddrStruct);

	char* ip;

    for (ifa = ifAddrStruct; ifa != NULL; ifa = ifa->ifa_next) {
        if (ifa ->ifa_addr->sa_family==AF_INET) { // check it is IP4
            // is a valid IP4 Address
            tmpAddrPtr=&((struct sockaddr_in *)ifa->ifa_addr)->sin_addr;
            char addressBuffer[INET_ADDRSTRLEN];
            inet_ntop(AF_INET, tmpAddrPtr, addressBuffer, INET_ADDRSTRLEN);
            printf("%s IP Address %s\n", ifa->ifa_name, addressBuffer); 
        } else if (ifa->ifa_addr->sa_family==AF_INET6) { // check it is IP6
            // is a valid IP6 Address
            tmpAddrPtr=&((struct sockaddr_in6 *)ifa->ifa_addr)->sin6_addr;
            char addressBuffer[INET6_ADDRSTRLEN];
            inet_ntop(AF_INET6, tmpAddrPtr, addressBuffer, INET6_ADDRSTRLEN);
            printf("%s IP Address %s\n", ifa->ifa_name, addressBuffer); 
        } 
    }
    if (ifAddrStruct!=NULL) freeifaddrs(ifAddrStruct);

}

int CConnect::connect_to_server() {
	int i;
	for( i = 0; i < 4; i++) {

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
  	sa.sin_port= htons(10333);                /* this is our port number */
  	udt_client = UDT::socket(AF_INET, SOCK_STREAM, 0); /* create socket */

  	if (UDT::bind(udt_client,(struct sockaddr *)&sa,sizeof(struct sockaddr_in)) == UDT::ERROR ) {
  		cout << "bind: " << UDT::getlasterror().getErrorMessage() << endl;
      		return 0;
	}

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

	cout<<"Connected to server";
	
	message report;
	report.type = REPORTIP;
	report.length = 50;

	char buffer[sizeof(message)];
	memcpy(buffer, &report, sizeof(message));
	int ss;

	if (UDT::ERROR == (ss = UDT::send(udt_client, buffer, sizeof(message), 0)))
	{
        cout << "send:" << UDT::getlasterror().getErrorMessage() << endl;
    }
	
	print_private_ip();
	
	UDT::close(udt_client);
	}
}

int main() {
	UDTUpDown _udt_;
	CConnect* connection = new CConnect();
	connection->connect_to_server();
	return 0;
}




