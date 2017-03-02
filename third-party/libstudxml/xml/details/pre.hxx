// file      : xml/details/pre.hxx
// copyright : Copyright (c) 2013-2017 Code Synthesis Tools CC
// license   : MIT; see accompanying LICENSE file

#ifdef _MSC_VER
  // Push warning state.
  //
#  pragma warning (push, 3)

  // Disabled warnings.
  //
#  pragma warning (disable:4251) // needs to have DLL-interface
#  pragma warning (disable:4290) // exception specification ignored
#  pragma warning (disable:4355) // passing 'this' to a member
#  pragma warning (disable:4800) // forcing value to bool

// Elevated warnings.
//
#  pragma warning (2:4239) // standard doesn't allow this conversion

#endif
