// file      : libstudxml/value-traits.txx
// copyright : Copyright (c) 2013-2019 Code Synthesis Tools CC
// license   : MIT; see accompanying LICENSE file

#include <sstream>

#include <libstudxml/parser.hxx>
#include <libstudxml/serializer.hxx>

namespace xml
{
  template <typename T>
  T default_value_traits<T>::
  parse (std::string s, const parser& p)
  {
    T r;
    std::istringstream is (s);
    if (!(is >> r && is.eof ()) )
      throw parsing (p, "invalid value '" + s + "'");
    return r;
  }

  template <typename T>
  std::string default_value_traits<T>::
  serialize (const T& v, const serializer& s)
  {
    std::ostringstream os;
    if (!(os << v))
      throw serialization (s, "invalid value");
    return os.str ();
  }
}
