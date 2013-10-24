#include <dis.hxx>
#include <stdio.h>

int main() 
{	
	char filepath[256];
	FILE * file = fopen("path_to_datafile.txt","r");

	if( file ) {
		fgets(filepath, sizeof(filepath), file);
		printf( "String gelesen: %s\n", filepath );
	}






	int DIM_data;
	DimService runNumber("gas_composition",DIM_data); 
	DimServer::start("DIM_Gaschromat"); 

	long int filesize;
	long int memorize_filesize = 0;
	FILE *fd;
	char filename[] = "file.dat";       // file to read

	while(1) 
	{ 
		// find size of file (in order to see if it has changed)
		FILE * fpfilesize = fopen(filename, "r");
		sleep(3);
		if (fpfilesize) {
			fseek(fpfilesize, 0, SEEK_END);
			filesize = ftell(fpfilesize);
			printf("filesize: %i\n",filesize);
			printf("memorize_filesize: %i\n",memorize_filesize);
			if (filesize > memorize_filesize) {
				memorize_filesize = filesize;
		
				// read last line of file only:
				static const long max_len = 55+ 1;  // define the max length of the line to read
				char buff[max_len + 1];             // define the buffer and allocate the length

				if ((fd = fopen(filename, "rb")) != NULL)  {      // open file. I omit error checks

					fseek(fd, -max_len, SEEK_END);            // set pointer to the end of file minus the length you need. Presumably there can be more than one new line caracter
					fread(buff, max_len-1, 1, fd);            // read the contents of the file starting from where fseek() positioned us
					fclose(fd);                               // close the file

					buff[max_len-1] = '\0';                   // close the string
					char *last_newline = strrchr(buff, '\n'); // find last occurrence of newlinw 
					char *last_line = last_newline+1;         // jump to it

					DIM_data = atoi(last_line);		  // convert last line of file (string) to int for DIM Server
										  // use atof() for float, atoi() for int
					runNumber.updateService(); 
				}
			}
			fseek(fpfilesize, 0, SEEK_SET);
			fclose(fpfilesize);
		}
	}
}

