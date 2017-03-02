// file      : xml/serializer.ixx
// copyright : Copyright (c) 2013-2017 Code Synthesis Tools CC
// license   : MIT; see accompanying LICENSE file

#include <xml/value-traits>

namespace xml
{
  // serialization
  //
  inline serialization::
  serialization (const std::string& name, const std::string& d)
      : name_ (name), description_ (d)
  {
    init ();
  }

  inline serialization::
  serialization (const serializer& s, const std::string& d)
      : name_ (s.output_name ()), description_ (d)
  {
    init ();
  }

  // serializer
  //
  inline void serializer::
  start_element (const qname_type& qname)
  {
    start_element (qname.namespace_ (), qname.name ());
  }

  inline void serializer::
  start_element (const std::string& name)
  {
    start_element (std::string (), name);
  }

  inline void serializer::
  end_element (const qname_type& qname)
  {
    end_element (qname.namespace_ (), qname.name ());
  }

  inline void serializer::
  end_element (const std::string& name)
  {
    end_element (std::string (), name);
  }

  inline void serializer::
  element (const std::string& v)
  {
    if (!v.empty ())
      characters (v);

    end_element ();
  }

  template <typename T>
  inline void serializer::
  element (const T& v)
  {
    element (value_traits<T>::serialize (v, *this));
  }

  inline void serializer::
  element (const std::string& n, const std::string& v)
  {
    element (std::string (), n, v);
  }

  template <typename T>
  inline void serializer::
  element (const std::string& n, const T& v)
  {
    element (n, value_traits<T>::serialize (v, *this));
  }

  inline void serializer::
  element (const qname_type& qn, const std::string& v)
  {
    element (qn.namespace_ (), qn.name (), v);
  }

  template <typename T>
  inline void serializer::
  element (const qname_type& qn, const T& v)
  {
    element (qn, value_traits<T>::serialize (v, *this));
  }

  template <typename T>
  inline void serializer::
  element (const std::string& ns, const std::string& n, const T& v)
  {
    element (ns, n, value_traits<T>::serialize (v, *this));
  }

  inline void serializer::
  start_attribute (const qname_type& qname)
  {
    start_attribute (qname.namespace_ (), qname.name ());
  }

  inline void serializer::
  start_attribute (const std::string& name)
  {
    start_attribute (std::string (), name);
  }

  inline void serializer::
  end_attribute (const qname_type& qname)
  {
    end_attribute (qname.namespace_ (), qname.name ());
  }

  inline void serializer::
  end_attribute (const std::string& name)
  {
    end_attribute (std::string (), name);
  }

  inline void serializer::
  attribute (const qname_type& qname, const std::string& value)
  {
    attribute (qname.namespace_ (), qname.name (), value);
  }

  template <typename T>
  inline void serializer::
  attribute (const qname_type& qname, const T& value)
  {
    attribute (qname, value_traits<T>::serialize (value, *this));
  }

  inline void serializer::
  attribute (const std::string& name, const std::string& value)
  {
    attribute (std::string (), name, value);
  }

  template <typename T>
  inline void serializer::
  attribute (const std::string& name, const T& value)
  {
    attribute (name, value_traits<T>::serialize (value, *this));
  }

  template <typename T>
  inline void serializer::
  attribute (const std::string& ns, const std::string& name, const T& value)
  {
    attribute (ns, name, value_traits<T>::serialize (value, *this));
  }

  template <typename T>
  inline void serializer::
  characters (const T& value)
  {
    characters (value_traits<T>::serialize (value, *this));
  }

  // operator<<
  //

  inline serializer&
  operator<< (serializer& s, void (*func) (serializer&))
  {
    func (s);
    return s;
  }

  namespace details
  {
    // Detect whether T is callable with argument A.
    //
    template <typename T, typename A>
    struct is_callable
    {
      typedef char no[1];
      typedef char yes[2];
      template <typename X> static X declval ();

      template <int> struct check;

      template <typename>
      static no& test (...);

      template <typename X>
      static yes& test (check<sizeof (declval<X> () (declval<A> ()), 0)>*);

      static const bool value = sizeof (test<T> (0)) == sizeof (yes);
    };

    template <typename T, bool = is_callable<const T&, serializer&>::value>
    struct inserter;

    template <typename T>
    struct inserter<T, true>
    {
      static void insert (serializer& s, const T& f) {f (s);}
    };

    template <typename T>
    struct inserter<T, false>
    {
      static void insert (serializer& s, const T& v) {s.characters (v);}
    };
  }

  template <typename T>
  inline serializer&
  operator<< (serializer& s, const T& value)
  {
    details::inserter<T>::insert (s, value);
    return s;
  }
}
