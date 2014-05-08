#pragma once

#include <opc/opc.h>

namespace xlnt {

class opc_callback_handler
{
public:
	static int read(void *context, char *buffer, int length);

	static int write(void *context, const char *buffer, int length);

	static int close(void *context);

	static opc_ofs_t seek(void *context, opc_ofs_t ofs);

	static int trim(void *context, opc_ofs_t new_size);

	static int flush(void *context);
};

} // namespace xlnt
