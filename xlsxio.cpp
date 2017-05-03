#ifdef _WINDOWS
#define WIN32_MEAN_AND_LEAN
#include <windows.h>
#endif

#include <lua.hpp>

#include "xlsxio_version.h"
#include "xlsxio_read.h"
#include "xlsxio_write.h"
#include "xlsxio.h"

fn_xlsxioread_open read_open = NULL;
fn_xlsxioread_close read_close = NULL;
fn_xlsxioread_sheetlist_open read_sheetlist_open = NULL;
fn_xlsxioread_sheetlist_close read_sheetlist_close = NULL;
fn_xlsxioread_sheetlist_next read_sheetlist_next = NULL;
fn_xlsxioread_process read_process = NULL;
fn_xlsxioread_sheet_open read_sheet_open = NULL;
fn_xlsxioread_sheet_close read_sheet_close = NULL;
fn_xlsxioread_sheet_next_row read_sheet_next_row = NULL;
fn_xlsxioread_sheet_next_cell read_sheet_next_cell = NULL;

fn_xlsxiowrite_open write_open = NULL;
fn_xlsxiowrite_close write_close = NULL;
fn_xlsxiowrite_set_detection_rows write_set_detection_rows = NULL;
fn_xlsxiowrite_set_row_height write_set_row_height = NULL;
fn_xlsxiowrite_add_column write_add_column = NULL;
fn_xlsxiowrite_add_cell_string write_add_cell_string = NULL;
fn_xlsxiowrite_add_cell_int write_add_cell_int = NULL;
fn_xlsxiowrite_add_cell_float write_add_cell_float = NULL;
fn_xlsxiowrite_add_cell_datetime write_add_cell_datetime = NULL;
fn_xlsxiowrite_next_row write_next_row = NULL;

struct ReaderData
{
	lua_State	*L;

	int			tableIndex;
	int			currentSheetTable, currentRow;

	int			cellCallbackIndex, rowCallbackIndex;
	int			endIterator;

	int			nowReadCols, startReadCol, maxReadCols;
	int			nowReadRows, startReadRow, maxReadRows;
	int			numOfEmptyCells;
	int			dropEmptyRow;
};

struct WriterData
{
	lua_State	*L;
};

int xlsxioread_process_cell_callback_totbl(size_t row, size_t col, const char* value, void* callbackdata)
{
	ReaderData* c = (ReaderData*)callbackdata;
	if (c->maxReadCols && c->nowReadCols >= c->maxReadCols)
		return 1;

	if (c->currentRow && c->nowReadCols >= c->startReadCol)
	{
		//value = (value && value[0]) ? translateGBK(value) : NULL;
		if (value && value[0])
			lua_pushstring(c->L, value);
		else
			lua_pushnil(c->L);
		lua_rawseti(c->L, c->currentRow, c->nowReadCols + 1);
	}

	c->nowReadCols ++;
	return 0;
}

int xlsxioread_process_row_callback_totbl(size_t row, size_t maxcol, void* callbackdata)
{
	ReaderData* c = (ReaderData*)callbackdata;

	if (c->currentRow)
	{
		if (!c->dropEmptyRow || c->numOfEmptyCells < c->nowReadCols)
			lua_rawseti(c->L, c->currentSheetTable, c->nowReadRows - c->startReadRow + 1);
		c->currentRow = 0;
	}

	c->nowReadRows ++;
	c->nowReadCols = 0;
	c->numOfEmptyCells = 0;

	if (c->maxReadRows && c->nowReadRows >= c->maxReadRows)
		return 1;

	if (c->nowReadRows >= c->startReadRow)
	{
		lua_newtable(c->L);
		c->currentRow = lua_gettop(c->L);
	}

	return 0;
}

static int doCallback(ReaderData* c, int top, int cbIndex)
{
	lua_State* L = c->L;
	lua_pushvalue(L, cbIndex);

	int r = lua_pcall(L, 1, 1, 0);
	if (r)
	{
		c->endIterator = 1;
		luaL_error(L, "Error when xlsreader callback: %s", lua_tostring(L, -1), 0);		
		return -1;
	}

	if (lua_gettop(L) > top)
	{
		if (lua_toboolean(L, -1) == 0)
		{
			c->endIterator = 1;
			return -1;
		}

		lua_settop(L, top);
	}

	return 0;
}

int xlsxioread_process_cell_callback_cb(size_t row, size_t col, const char* value, void* callbackdata)
{
	ReaderData* c = (ReaderData*)callbackdata;
	if (c->maxReadCols && c->nowReadCols >= c->maxReadCols)
		return 1;

	if (c->currentRow && c->nowReadCols >= c->startReadCol && c->cellCallbackIndex)
	{
		if (value && value[0])
			lua_pushstring(c->L, value);
		else
			lua_pushnil(c->L);

		if (doCallback(c, lua_gettop(c->L) - 1, c->cellCallbackIndex))
			return -1;
	}

	c->nowReadCols ++;
	return 0;
}

int xlsxioread_process_row_callback_cb(size_t row, size_t maxcol, void* callbackdata)
{
	ReaderData* c = (ReaderData*)callbackdata;

	if (c->currentRow)
	{
		if (!c->dropEmptyRow || c->numOfEmptyCells < c->nowReadCols)
			lua_rawseti(c->L, c->currentSheetTable, c->nowReadRows - c->startReadRow + 1);
		c->currentRow = 0;
	}

	c->nowReadRows ++;
	c->nowReadCols = 0;
	c->numOfEmptyCells = 0;

	if (c->maxReadRows && c->nowReadRows >= c->maxReadRows)
		return 1;

	if (c->nowReadRows >= c->startReadRow && c->rowCallbackIndex)
	{
		if (doCallback(c, lua_gettop(c->L), c->rowCallbackIndex))
			return -1;
	}

	return 0;
}

static int lua_xlsx_read(lua_State* L)
{
	const char* filename = luaL_checkstring(L, 1);

	int nn = 2;
	ReaderData ud;
	memset(&ud, 0, sizeof(ud));
	ud.L = L;

	if (lua_istable(L, 2))
	{
		lua_pushliteral(L, "dropemptyrow");
		lua_rawget(L, 2);
		ud.dropEmptyRow = lua_toboolean(L, -1);

		lua_pushliteral(L, "maxrows");
		lua_rawget(L, 2);
		if (lua_isnumber(L, -1))
			ud.maxReadRows = lua_tointeger(L, -1);

		lua_pushliteral(L, "maxcols");
		lua_rawget(L, 2);
		if (lua_isnumber(L, -1))
			ud.maxReadCols = lua_tointeger(L, -1);

		lua_pushliteral(L, "startrow");
		lua_rawget(L, 2);
		if (lua_isnumber(L, -1))
			ud.startReadRow = lua_tointeger(L, -1);

		lua_pushliteral(L, "startcol");
		lua_rawget(L, 2);
		if (lua_isnumber(L, -1))
			ud.startReadCol = lua_tointeger(L, -1);

		lua_pushliteral(L, "totable");
		lua_rawget(L, 2);
		if (lua_toboolean(L, -1))
		{
			lua_newtable(L);
			ud.tableIndex = lua_gettop(L);
		}

		nn ++;
	}

	if (!ud.tableIndex)
	{
		// 两个回调，第一个是cell的，第二个是row的
		int cb1 = lua_type(L, nn);
		int cb2 = lua_type(L, nn + 1);

		if (cb1 == LUA_TFUNCTION)
			ud.cellCallbackIndex = nn;
		else if (cb1 != LUA_TNIL && cb1 != LUA_TNONE)
			luaL_error(L, "The third parameter for xlsxio.read function must be function (or nil) type", 0);

		if (cb2 == LUA_TFUNCTION)
			ud.rowCallbackIndex = nn;
		else if (cb2 != LUA_TNIL && cb2 != LUA_TNONE)
			luaL_error(L, "The forth parameter for xlsxio.read function must be function (or nil) type", 0);

		if (cb1 != LUA_TFUNCTION && cb2 != LUA_TFUNCTION)
		{
			luaL_error(L, "Any callback function in third (or forth) parameter", 0);
			return 0;
		}
	}

	void* reader = read_open(filename);
	if (!reader)
	{
		lua_pushboolean(L, 0);
		return 1;
	}

	void* list = read_sheetlist_open(reader);
	for (int iSheet = 0; ; ++ iSheet)
	{
		char* name = read_sheetlist_next(list);
		if (!name)
			break;
		
		if (ud.tableIndex)
		{
			// 转Table模式
			lua_newtable(L);

				lua_pushliteral(L, "name");
				lua_pushstring(L, name);
				lua_rawset(L, -3);

				lua_pushliteral(L, "rows");
				lua_newtable(L);
				ud.currentSheetTable = lua_gettop(L);

					lua_newtable(L);
					ud.currentRow = lua_gettop(L);

					read_process(reader, name, 0, &xlsxioread_process_cell_callback_totbl, &xlsxioread_process_row_callback_totbl, &ud);

				if (ud.currentRow)
				{
					lua_pop(L, 1);
					ud.currentRow = 0;
				}

				lua_rawset(L, -3);

			lua_rawseti(L, ud.tableIndex, iSheet + 1);
		}
		else
		{
			// 回调模式
			read_process(reader, name, 0, &xlsxioread_process_cell_callback_cb, &xlsxioread_process_row_callback_cb, &ud);
			if (ud.endIterator)
				break;
		}

		ud.currentSheetTable = 0;
		ud.nowReadCols = ud.nowReadRows = 0;
		ud.numOfEmptyCells = 0;
	}

	read_close(reader);

	if (!ud.tableIndex)
		lua_pushboolean(L, 1);

	return 1;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#define D2I_ENDIALOC	0
union double2int
{
	double	dval;
	int		ivals[2];
};

static int readRow(lua_State* L, void* writer, int tableIndex)
{
	int ival;
	double dval;
	double2int d2i;

	// 读配置
	lua_pushliteral(L, "height");
	lua_rawget(L, tableIndex);
	if (lua_isnumber(L, -1))
		write_set_row_height(writer, lua_tointeger(L, -1));

	// 读列
	lua_pushliteral(L, "cols");
	lua_rawget(L, tableIndex);
	int rowsIndex = lua_gettop(L);
	if (lua_istable(L, rowsIndex))
	{
		lua_pushnil(L);
		while(lua_next(L, rowsIndex))
		{
			if (lua_isnumber(L, -2))
			{
				int tp = lua_type(L, -1);
				switch (tp)
				{
				case LUA_TNUMBER:
					dval = lua_tonumber(L, -1);
					d2i.dval = dval + 6755399441055744.0;
					ival = d2i.ivals[D2I_ENDIALOC];

					if ((double)ival == dval && abs(ival) < INT_MAX)
						write_add_cell_int(writer, d2i.ivals[D2I_ENDIALOC]);
					else
						write_add_cell_float(writer, dval);
					break;

				case LUA_TSTRING:
					write_add_cell_string(writer, lua_tostring(L, -1));
					break;
				}
			}

			lua_pop(L, 1);
		}
	}

	write_next_row(writer);

	return 0;
}

static int lua_xlsx_write(lua_State* L)
{
	int nn = 3, top;
	const char* filename = luaL_checkstring(L, 1);
	const char* sheetname = luaL_checkstring(L, 2);
	
	void* writer = write_open(filename, sheetname);
	if (!writer)
	{
		lua_pushboolean(L, 0);
		return 1;
	}

	if (lua_istable(L, nn))
	{
		lua_pushliteral(L, "detectrows");
		lua_rawget(L, nn);
		if (lua_isnumber(L, -1))
			write_set_detection_rows(writer, lua_tointeger(L, -1));

		lua_pushliteral(L, "headers");
		lua_rawget(L, nn);
		if (lua_istable(L, -1))
		{
			// 添加表头
			top = lua_gettop(L);
			for(int iCol = 1; ; iCol ++)
			{
				lua_rawgeti(L, top, iCol);
				if (lua_istable(L, -1))
				{
					// 表头的每一个都是一个Table，含有text+width两个成员
					lua_pushliteral(L, "text");
					lua_rawget(L, -2);
					const char* name = lua_tostring(L, -1);

					lua_pushliteral(L, "width");
					lua_rawget(L, -3);
					if (lua_isnumber(L, -1))
						write_add_column(writer, name, lua_tointeger(L, -1));
					else
						write_add_column(writer, name, 100);
				}
				else if (lua_isstring(L, -1))
				{
					// 表头的每一个就是字符串，宽度用默认的
					write_add_column(writer, lua_tostring(L, -1), 100);
				}
				else
					break;

				lua_settop(L, top);
			}

			write_next_row(writer);
		}

		nn ++;
	}

	int type = lua_type(L, nn);
	if (type == LUA_TTABLE)
	{
		top = lua_gettop(L);
		for (int iRow = 1; ; iRow ++)
		{
			lua_rawgeti(L, nn, iRow);
			if (!lua_istable(L, top + 1))
				break;
				
			readRow(L, writer, top + 1);
			lua_settop(L, top);
		}
	}

	write_close(writer);
	return 0;
}

struct InitLib
{
	HLIB hReadLib, hWriteLib;
	InitLib()
	{		
		hReadLib = loadLibXLSIOReader();
		hWriteLib = loadLibXLSIOWriter();
	}
	~InitLib()
	{
		if (hReadLib)
			dlclose(hReadLib);
		if (hWriteLib)
			dlclose(hWriteLib);
	}
} _g_InitLib;

#ifdef _WINDOWS
extern "C" __declspec(dllexport) int luaopen_xlsxio(lua_State* L)
#else
extern "C" __declspec(dllexport) int luaopen_libxlsxio(lua_State* L)
#endif
{
	if (!_g_InitLib.hReadLib)
	{
		lua_pushliteral(L, "load \"xlsxio_read\" dynamic library failed");
		return 1;
	}

	lua_newtable(L);

	lua_pushliteral(L, "read");
	lua_pushcfunction(L, &lua_xlsx_read);
	lua_rawset(L, -3);

	lua_pushliteral(L, "write");
	lua_pushcfunction(L, &lua_xlsx_write);
	lua_rawset(L, -3);

	return 1;
}