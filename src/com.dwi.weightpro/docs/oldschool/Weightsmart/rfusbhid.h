extern "C"
{

	int __stdcall rdy_read(unsigned char MemBank, unsigned char WordAdd, unsigned char WordCnt, unsigned char *PassWord,unsigned char 	*TagCount, unsigned char *DataLen, unsigned char *Data, unsigned char *ReadLen, unsigned char *AntID, unsigned char *ReadCount);
	int __stdcall rdy_write(unsigned char *PassWord, unsigned char MemBank, 
	       unsigned char WordAdd, unsigned char WordCnt, unsigned char *Data);
	int __stdcall rdy_set_access_epc_match(unsigned char Mode, unsigned char EpcLen, unsigned char *Epc);
	int __stdcall get_firmware_version(unsigned char *Version); 
	int __stdcall SetOutPower(unsigned char Value);
	int __stdcall GetOutPower(unsigned char *Value);

}