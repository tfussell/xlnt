// file      : xml/parser.txx
// copyright : Copyright (c) 2013-2017 Code Synthesis Tools CC
// license   : MIT; see accompanying LICENSE file

#include <xml/value-traits>

namespace xml
{
  template <typename T>
  T parser::
  attribute (const qname_type& qn, const T& dv) const
  {
    if (const element_entry* e = get_element ())
    {
      attribute_map_type::const_iterator i (e->attr_map_.find (qn));

      if (i != e->attr_map_.end ())
      {
        if (!i->second.handled)
        {
          i->second.handled = true;
          e->attr_unhandled_--;
        }
        return value_traits<T>::parse (i->second.value, *this);
      }
    }

    return dv;
  }

  template <typename T>
  T parser::
  element (const qname_type& qn, const T& dv)
  {
    if (peek () == start_element && qname () == qn)
    {
      next ();
      return element<T> ();
    }

    return dv;
  }
}
