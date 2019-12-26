// file      : libstudxml/content.hxx -*- C++ -*-
// copyright : Copyright (c) 2013-2019 Code Synthesis Tools CC
// license   : MIT; see accompanying LICENSE file

#ifndef LIBSTUDXML_CONTENT_HXX
#define LIBSTUDXML_CONTENT_HXX

#include <libstudxml/details/pre.hxx>

namespace xml
{
  // XML content model. C++11 enum class emulated for C++98.
  //
  struct content
  {
    enum value
    {
               //  element   characters  whitespaces        notes
      empty,   //    no          no        ignored
      simple,  //    no          yes       preserved   content accumulated
      complex, //    yes         no        ignored
      mixed    //    yes         yes       preserved
    };

    content (value v): v_ (v) {};
    operator value () const {return v_;}

  private:
    value v_;
  };
}

#include <libstudxml/details/post.hxx>

#endif // LIBSTUDXML_CONTENT_HXX
