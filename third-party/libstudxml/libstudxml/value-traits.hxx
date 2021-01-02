// file      : libstudxml/value-traits.hxx -*- C++ -*-
// copyright : Copyright (c) 2013-2019 Code Synthesis Tools CC
// license   : MIT; see accompanying LICENSE file

#ifndef LIBSTUDXML_VALUE_TRAITS_HXX
#define LIBSTUDXML_VALUE_TRAITS_HXX

#include <libstudxml/details/pre.hxx>

#include <string>
#include <cstddef> // std::size_t

#include <libstudxml/forward.hxx>

#include <libstudxml/details/export.hxx>

namespace xml
{
  template <typename T>
  struct default_value_traits
  {
    static T
    parse (std::string, const parser&);

    static std::string
    serialize (const T&, const serializer&);
  };

  template <>
  struct LIBSTUDXML_EXPORT default_value_traits<bool>
  {
    static bool
    parse (std::string, const parser&);

    static std::string
    serialize (bool v, const serializer&)
    {
      return v ? "true" : "false";
    }
  };

  template <>
  struct default_value_traits<std::string>
  {
    static std::string
    parse (std::string s, const parser&)
    {
      return s;
    }

    static std::string
    serialize (const std::string& v, const serializer&)
    {
      return v;
    }
  };

  template <typename T>
  struct value_traits: default_value_traits<T> {};

  template <typename T, std::size_t N>
  struct value_traits<T[N]>: default_value_traits<const T*> {};
}

#include <libstudxml/value-traits.txx>

#include <libstudxml/details/post.hxx>

#endif // LIBSTUDXML_VALUE_TRAITS_HXX
