// file      : xml/parser.ixx
// copyright : Copyright (c) 2013-2017 Code Synthesis Tools CC
// license   : MIT; see accompanying LICENSE file

#include <cassert>

#include <xml/value-traits>

namespace xml
{
  // parsing
  //
  inline parsing::
  parsing (const std::string& n,
           unsigned long long l,
           unsigned long long c,
           const std::string& d)
      : name_ (n), line_ (l), column_ (c), description_ (d)
  {
    init ();
  }

  inline parsing::
  parsing (const parser& p, const std::string& d)
      : name_ (p.input_name ()),
        line_ (p.line ()),
        column_ (p.column ()),
        description_ (d)
  {
    init ();
  }

  // parser
  //
  inline parser::
  parser (std::istream& is, const std::string& iname, feature_type f)
      : size_ (0), iname_ (iname), feature_ (f)
  {
    data_.is = &is;
    init ();
  }

  inline parser::
  parser (const void* data,
          std::size_t size,
          const std::string& iname,
          feature_type f)
      : size_ (size), iname_ (iname), feature_ (f)
  {
    assert (data != nullptr && size != 0);

    data_.buf = data;
    init ();
  }

  inline parser::event_type parser::
  peek ()
  {
    if (state_ == state_peek)
      return event_;
    else
    {
      event_type e (next_ (true));
      state_ = state_peek; // Set it after the call to next_().
      return e;
    }
  }

  template <typename T>
  inline T parser::
  value () const
  {
    return value_traits<T>::parse (value (), *this);
  }

  inline const parser::element_entry* parser::
  get_element () const
  {
    return element_state_.empty () ? nullptr : get_element_ ();
  }

  inline const std::string& parser::
  attribute (const std::string& n) const
  {
    return attribute (qname_type (n));
  }

  template <typename T>
  inline T parser::
  attribute (const std::string& n) const
  {
    return attribute<T> (qname_type (n));
  }

  inline std::string parser::
  attribute (const std::string& n, const std::string& dv) const
  {
    return attribute (qname_type (n), dv);
  }

  template <typename T>
  inline T parser::
  attribute (const std::string& n, const T& dv) const
  {
    return attribute<T> (qname_type (n), dv);
  }

  template <typename T>
  inline T parser::
  attribute (const qname_type& qn) const
  {
    return value_traits<T>::parse (attribute (qn), *this);
  }

  inline bool parser::
  attribute_present (const std::string& n) const
  {
    return attribute_present (qname_type (n));
  }

  inline const parser::attribute_map_type& parser::
  attribute_map () const
  {
    if (const element_entry* e = get_element ())
    {
      e->attr_unhandled_ = 0; // Assume all handled.
      return e->attr_map_;
    }

    return empty_attr_map_;
  }

  inline void parser::
  next_expect (event_type e, const qname_type& qn)
  {
    next_expect (e, qn.namespace_ (), qn.name ());
  }

  inline void parser::
  next_expect (event_type e, const std::string& n)
  {
    next_expect (e, std::string (), n);
  }

  template <typename T>
  inline T parser::
  element ()
  {
    return value_traits<T>::parse (element (), *this);
  }

  inline std::string parser::
  element (const std::string& n)
  {
    next_expect (start_element, n);
    return element ();
  }

  inline std::string parser::
  element (const qname_type& qn)
  {
    next_expect (start_element, qn);
    return element ();
  }

  template <typename T>
  inline T parser::
  element (const std::string& n)
  {
    return value_traits<T>::parse (element (n), *this);
  }

  template <typename T>
  inline T parser::
  element (const qname_type& qn)
  {
    return value_traits<T>::parse (element (qn), *this);
  }

  inline std::string parser::
  element (const std::string& n, const std::string& dv)
  {
    return element (qname_type (n), dv);
  }

  template <typename T>
  inline T parser::
  element (const std::string& n, const T& dv)
  {
    return element<T> (qname_type (n), dv);
  }

  inline void parser::
  content (content_type c)
  {
    assert (state_ == state_next);

    if (!element_state_.empty () && element_state_.back ().depth == depth_)
      element_state_.back ().content = c;
    else
      element_state_.push_back (element_entry (depth_, c));
  }

  inline parser::content_type parser::
  content () const
  {
    assert (state_ == state_next);

    return
      !element_state_.empty () && element_state_.back ().depth == depth_
      ? element_state_.back ().content
      : content_type (content_type::mixed);
  }

  inline void parser::
  next_expect (event_type e, const qname_type& qn, content_type c)
  {
    next_expect (e, qn);
    assert (e == start_element);
    content (c);
  }

  inline void parser::
  next_expect (event_type e, const std::string& n, content_type c)
  {
    next_expect (e, std::string (), n);
    assert (e == start_element);
    content (c);
  }

  inline void parser::
  next_expect (event_type e,
               const std::string& ns, const std::string& n,
               content_type c)
  {
    next_expect (e, ns, n);
    assert (e == start_element);
    content (c);
  }
}
