#ifdef WIN32
typedef HMODULE HLIB;
#	define dlsym GetProcAddress
#	define dlclose FreeLibrary
#else
typedef void* HLIB;
#endif

///
typedef void* (*fn_xlsxioread_open)(const char* filename);
typedef void(*fn_xlsxioread_close)(void* handle);

typedef void* (*fn_xlsxioread_sheetlist_open)(void* handle);
typedef void(*fn_xlsxioread_sheetlist_close)(void* sheetlisthandle);
typedef char* (*fn_xlsxioread_sheetlist_next)(void* sheetlisthandle);

typedef int(*fn_xlsxioread_process)(void* handle, const char* sheetname, unsigned int flags, xlsxioread_process_cell_callback_fn cell_callback, xlsxioread_process_row_callback_fn row_callback, void* callbackdata);

typedef void* (*fn_xlsxioread_sheet_open)(void* handle, const char* sheetname, unsigned int flags);
typedef void(*fn_xlsxioread_sheet_close)(void* sheethandle);
typedef int(*fn_xlsxioread_sheet_next_row)(void* sheethandle);
typedef char* (*fn_xlsxioread_sheet_next_cell)(void* sheethandle);

///
typedef void* (*fn_xlsxiowrite_open)(const char* filename, const char* sheetname);
typedef int (*fn_xlsxiowrite_close)(void* handle);
typedef void (*fn_xlsxiowrite_set_detection_rows)(void* handle, size_t rows);
typedef void (*fn_xlsxiowrite_set_row_height)(void* handle, size_t height);
typedef void (*fn_xlsxiowrite_add_column)(void* handle, const char* name, int width);
typedef void (*fn_xlsxiowrite_add_cell_string)(void* handle, const char* value);
typedef void (*fn_xlsxiowrite_add_cell_int)(void* handle, int64_t value);
typedef void (*fn_xlsxiowrite_add_cell_float)(void* handle, double value);
typedef void (*fn_xlsxiowrite_add_cell_datetime)(void* handle, time_t value);
typedef void (*fn_xlsxiowrite_next_row)(void* handle);

///
extern fn_xlsxioread_open read_open;
extern fn_xlsxioread_close read_close;
extern fn_xlsxioread_sheetlist_open read_sheetlist_open;
extern fn_xlsxioread_sheetlist_close read_sheetlist_close;
extern fn_xlsxioread_sheetlist_next read_sheetlist_next;
extern fn_xlsxioread_process read_process;
extern fn_xlsxioread_sheet_open read_sheet_open;
extern fn_xlsxioread_sheet_close read_sheet_close;
extern fn_xlsxioread_sheet_next_row read_sheet_next_row;
extern fn_xlsxioread_sheet_next_cell read_sheet_next_cell;

extern fn_xlsxiowrite_open write_open;
extern fn_xlsxiowrite_close write_close;
extern fn_xlsxiowrite_set_detection_rows write_set_detection_rows;
extern fn_xlsxiowrite_set_row_height write_set_row_height;
extern fn_xlsxiowrite_add_column write_add_column;
extern fn_xlsxiowrite_add_cell_string write_add_cell_string;
extern fn_xlsxiowrite_add_cell_int write_add_cell_int;
extern fn_xlsxiowrite_add_cell_float write_add_cell_float;
extern fn_xlsxiowrite_add_cell_datetime write_add_cell_datetime;
extern fn_xlsxiowrite_next_row write_next_row;

///
static HLIB loadLibXLSIOReader()
{
#ifdef _WINDOWS
	HLIB h = LoadLibrary(L"libxlsxio_read.dll");
#else
	HLIB h = dlopen("libxlsxio_read", RTLD_LAZY);
#endif
	if (!h)
		return NULL;

	read_open = (fn_xlsxioread_open)dlsym(h, "xlsxioread_open");
	read_close = (fn_xlsxioread_close)dlsym(h, "xlsxioread_close");
	read_sheetlist_open = (fn_xlsxioread_sheetlist_open)dlsym(h, "xlsxioread_sheetlist_open");
	read_sheetlist_close = (fn_xlsxioread_sheetlist_close)dlsym(h, "xlsxioread_sheetlist_close");
	read_sheetlist_next = (fn_xlsxioread_sheetlist_next)dlsym(h, "xlsxioread_sheetlist_next");
	read_process = (fn_xlsxioread_process)dlsym(h, "xlsxioread_process");
	read_sheet_open = (fn_xlsxioread_sheet_open)dlsym(h, "xlsxioread_sheet_open");
	read_sheet_close = (fn_xlsxioread_sheet_close)dlsym(h, "xlsxioread_sheet_close");
	read_sheet_next_row = (fn_xlsxioread_sheet_next_row)dlsym(h, "xlsxioread_sheet_next_row");
	read_sheet_next_cell = (fn_xlsxioread_sheet_next_cell)dlsym(h, "xlsxioread_sheet_next_cell");

	return h;
}

static HLIB loadLibXLSIOWriter()
{
#ifdef _WINDOWS
	HLIB h = LoadLibrary(L"libxlsxio_write.dll");
#else
	HLIB h = dlopen("libxlsxio_write", RTLD_LAZY);
#endif
	if (!h)
		return NULL;

	write_open = (fn_xlsxiowrite_open)dlsym(h, "xlsxiowrite_open");
	write_close = (fn_xlsxiowrite_close)dlsym(h, "xlsxiowrite_close");
	write_set_detection_rows = (fn_xlsxiowrite_set_detection_rows)dlsym(h, "xlsxiowrite_set_detection_rows");
	write_set_row_height = (fn_xlsxiowrite_set_row_height)dlsym(h, "xlsxiowrite_set_row_height");
	write_add_column = (fn_xlsxiowrite_add_column)dlsym(h, "xlsxiowrite_add_column");
	write_add_cell_string = (fn_xlsxiowrite_add_cell_string)dlsym(h, "xlsxiowrite_add_cell_string");
	write_add_cell_int = (fn_xlsxiowrite_add_cell_int)dlsym(h, "xlsxiowrite_add_cell_int");
	write_add_cell_float = (fn_xlsxiowrite_add_cell_float)dlsym(h, "xlsxiowrite_add_cell_float");
	write_add_cell_datetime = (fn_xlsxiowrite_add_cell_datetime)dlsym(h, "xlsxiowrite_add_cell_datetime");
	write_next_row = (fn_xlsxiowrite_next_row)dlsym(h, "xlsxiowrite_next_row");

	return h;
}