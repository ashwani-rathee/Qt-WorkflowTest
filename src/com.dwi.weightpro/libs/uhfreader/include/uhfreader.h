extern "C"
{

int __stdcall  OpenCom(__int16 port,long baud);
int __stdcall  CloseCom();
int  __stdcall Net_Connect(char *HostIp,int HostPort,
			char *ReaderIp,int ReaderPort);
int __stdcall  RSSIToLog(unsigned char RSSI);

int __stdcall SetWorkAntenna(unsigned char AntennaID);
int __stdcall GetWorkAntenna(unsigned char *AntennaID);
int __stdcall SetDRMStatus(unsigned char Status);
int __stdcall GetDRMStatus(unsigned char *Status);
int __stdcall SetOutPower(unsigned char Value);
int __stdcall GetOutPower(unsigned char *Value);
int __stdcall SetRS485Address(unsigned char Addr);
int __stdcall GetGpioStatus(unsigned char *Gpio1Status,unsigned char *Gpio2Status);
int __stdcall SetGpioStatus(unsigned char ChooseGpio,unsigned char GpioValue);
int __stdcall SetFrequencyRegion(unsigned char Region,
					int StartFreq,int EndFreq,int FreqSpace,
					unsigned char FreqQuantity);
int __stdcall GetFrequencyRegion(unsigned char *Region,
					int *StartFreq,int *EndFreq,int *FreqSpace,
					unsigned char *FreqQuantity);
int __stdcall SetBeeperMode(unsigned char ModeValue);
int __stdcall SetBeeper(unsigned char Value);
int __stdcall GetReaderTemperature(unsigned char *PlusMinus,unsigned char *TempValue);
int __stdcall SetAntConnectionDetector(unsigned char Sensitivity);
int __stdcall GetAntConnectionDetector(unsigned char *Sensitivity);
int __stdcall GetTidFastStatus(unsigned char *Status);
int __stdcall SetTidFastStatus(unsigned char Status);
int __stdcall GetRfComLinkStatus(unsigned char *Status);
int __stdcall SetRfComLinkStatus(unsigned char Status);
int __stdcall Reset();
int __stdcall SetUartBaudRate(unsigned char BaudRate);
int __stdcall GetFirmwareVersion(unsigned char *Version);
int __stdcall GetRfPortReturnLoss(unsigned char FreqParameter,unsigned char *LossValue);
int __stdcall SetReaderIdentifier(unsigned char *IdValue);
int __stdcall GetReaderIdentifier(unsigned char *IdValue);

int __stdcall ReadTag(unsigned char MemBank, unsigned char WordAdd, unsigned char WordCnt, unsigned char *PassWord,
						unsigned char *TagCount, unsigned char *DataLen, unsigned char *Data, unsigned char *ReadLen, 
						unsigned char *AntID, unsigned char *ReadCount);
int __stdcall WriteTag(unsigned char *PassWord, unsigned char MemBank,
	unsigned char WordAdd, unsigned char WordCnt, unsigned char *Data);

int __stdcall Testinventory(unsigned char Repeat,unsigned char * TagCount,unsigned char * DataLen);

int __stdcall Inventory(unsigned char Repeat,unsigned char *OutData,unsigned char *TagNum);
int __stdcall CleanInventory();

int __stdcall SetAccessEpcMatch(unsigned char Mode, unsigned char EpcLen, unsigned char *EpcData);


}
