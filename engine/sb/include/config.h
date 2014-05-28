
#ifndef __CONFIG_H__
#define __CONFIG_H__ 1

typedef struct _CONFIG{
	char* module;// "c:\\scriptbasic\\modules\\"
	char* include; // "c:\\scriptbasic\\include\\"
	char* cache;// "c:\\scriptbasic\\cache\\"
	int maxinclude; // 100
	int maxstep; // 0
	int maxlocalstep;// 0
	int maxlevel;// 3000
	int maxmem; // 0
	int maxderef;
}CONFIG, *psCONFIG;

#endif
