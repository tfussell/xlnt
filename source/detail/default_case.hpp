#pragma once

#define EXCEPT_ON_UNHANDLED_SWITCH_CASE

#ifdef EXCEPT_ON_UNHANDLED_SWITCH_CASE
#define default_case(default_value) throw xlnt::exception("unhandled case");
#else
#define default_case(default_value) return default_value;
#endif
