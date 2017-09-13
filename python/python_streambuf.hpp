#pragma once

#include <cassert>
#include <stdexcept>
#include <iostream>
#include <pybind11/pybind11.h>

namespace xlnt {

class python_streambuf : public std::basic_streambuf<char>
{
  private:
    typedef std::basic_streambuf<char> base_t;

  public:
    /* The syntax
        using base_t::char_type;
       would be nicer but Visual Studio C++ 8 chokes on it
    */
    typedef base_t::char_type   char_type;
    typedef base_t::int_type    int_type;
    typedef base_t::pos_type    pos_type;
    typedef base_t::off_type    off_type;
    typedef base_t::traits_type traits_type;

    // work around Visual C++ 7.1 problem
    inline static int traits_type_eof()
    {
        return traits_type::eof();
    }

    /// The default size of the read and write buffer.
    /** They are respectively used to buffer data read from and data written to
        the Python file object. It can be modified from Python.
    */
    static std::size_t default_buffer_size;

    /// Construct from a Python file object
    /** if buffer_size is 0 the current default_buffer_size is used.
    */
    python_streambuf(
      pybind11::object python_file_obj,
      std::size_t buffer_size_ = 0)
    :
      py_read(python_file_obj.attr("read").cast<pybind11::function>()),
      py_write(python_file_obj.attr("write").cast<pybind11::function>()),
      py_seek(python_file_obj.attr("seek").cast<pybind11::function>()),
      py_tell(python_file_obj.attr("tell").cast<pybind11::function>()),
      buffer_size(buffer_size_ != 0 ? buffer_size_ : default_buffer_size),
      write_buffer(0),
      pos_of_read_buffer_end_in_py_file(0),
      pos_of_write_buffer_end_in_py_file(buffer_size),
      farthest_pptr(0)
    {
      assert(buffer_size != 0);

      /* Some Python file objects (e.g. sys.stdout and sys.stdin)
         have non-functional seek and tell. If so, assign None to
         py_tell and py_seek.
       */
       if (!py_tell.is_none())
       {
         try
         {
           py_tell();
         }
         catch(...)
         {
           py_tell = pybind11::none();
         }
       }

      if (!py_write.is_none())
      {
        // C-like string to make debugging easier
        write_buffer = new char[buffer_size + 1];
        write_buffer[buffer_size] = '\0';
        setp(write_buffer, write_buffer + buffer_size);  // 27.5.2.4.5 (5)
        farthest_pptr = pptr();
      }
      else
      {
        // The first attempt at output will result in a call to overflow
        setp(0, 0);
      }

      if (!py_tell.is_none())
      {
        auto py_pos = py_tell().cast<pybind11::int_>();
        pos_of_read_buffer_end_in_py_file = py_pos;
        pos_of_write_buffer_end_in_py_file = py_pos;
      }
    }

    /// Mundane destructor freeing the allocated resources
    virtual ~python_streambuf() {
      if (write_buffer) delete[] write_buffer;
    }

    /// C.f. C++ standard section 27.5.2.4.3
    /** It is essential to override this virtual function for the stream
        member function readsome to work correctly (c.f. 27.6.1.3, alinea 30)
     */
    virtual std::streamsize showmanyc() {
      int_type const failure = traits_type::eof();
      int_type status = underflow();
      if (status == failure) return -1;
      return egptr() - gptr();
    }

    /// C.f. C++ standard section 27.5.2.4.3
    virtual int_type underflow() {
      int_type const failure = traits_type::eof();
      if (py_read.is_none()) {
        throw std::invalid_argument(
          "That Python file object has no 'read' attribute");
      }
      read_buffer = py_read(buffer_size).cast<pybind11::bytes>();
      char *read_buffer_data = nullptr;
      Py_ssize_t py_n_read = 0;
      if (PyBytes_AsStringAndSize(read_buffer.ptr(), &read_buffer_data, &py_n_read) == -1) {
        setg(0, 0, 0);
        throw std::invalid_argument(
          "The method 'read' of the Python file object "
          "did not return a string.");
      }
      auto n_read = (off_type)py_n_read;
      pos_of_read_buffer_end_in_py_file += n_read;
      setg(read_buffer_data, read_buffer_data, read_buffer_data + n_read);
      // ^^^27.5.2.3.1 (4)
      if (n_read == 0) return failure;
      return traits_type::to_int_type(read_buffer_data[0]);
    }

    /// C.f. C++ standard section 27.5.2.4.5
    virtual int_type overflow(int_type c=traits_type_eof()) {
      if (py_write.is_none()) {
        throw std::invalid_argument(
          "That Python file object has no 'write' attribute");
      }
      farthest_pptr = std::max(farthest_pptr, pptr());
      auto n_written = (off_type)(farthest_pptr - pbase());
      auto chunk = PyBytes_FromStringAndSize(pbase(), farthest_pptr - pbase());
      py_write(chunk);
      if (!traits_type::eq_int_type(c, traits_type::eof())) {
	      auto ch = traits_type::to_char_type(c);
        py_write(reinterpret_cast<char *>(&ch), 1);
        n_written++;
      }
      if (n_written) {
        pos_of_write_buffer_end_in_py_file += n_written;
        setp(pbase(), epptr());
        // ^^^ 27.5.2.4.5 (5)
        farthest_pptr = pptr();
      }
      return traits_type::eq_int_type(
        c, traits_type::eof()) ? traits_type::not_eof(c) : c;
    }

    /// Update the python file to reflect the state of this stream buffer
    /** Empty the write buffer into the Python file object and set the seek
        position of the latter accordingly (C++ standard section 27.5.2.4.2).
        If there is no write buffer or it is empty, but there is a non-empty
        read buffer, set the Python file object seek position to the
        seek position in that read buffer.
    */
    virtual int sync() {
      int result = 0;
      farthest_pptr = std::max(farthest_pptr, pptr());
      if (farthest_pptr && farthest_pptr > pbase()) {
        off_type delta = pptr() - farthest_pptr;
        int_type status = overflow();
        if (traits_type::eq_int_type(status, traits_type::eof())) result = -1;
        if (!py_seek.is_none())
        {
          py_seek(delta);
        }
      }
      else if (gptr() && gptr() < egptr()) {
        if (!py_seek.is_none())
        {
          py_seek(gptr() - egptr(), 1);
        }
      }
      return result;
    }

    /// C.f. C++ standard section 27.5.2.4.2
    /** This implementation is optimised to look whether the position is within
        the buffers, so as to avoid calling Python seek or tell. It is
        important for many applications that the overhead of calling into Python
        is avoided as much as possible (e.g. parsers which may do a lot of
        backtracking)
    */
    virtual
    pos_type seekoff(off_type off, std::ios_base::seekdir way,
                     std::ios_base::openmode which=  std::ios_base::in
                                                   | std::ios_base::out)
    {
      /* In practice, "which" is either std::ios_base::in or out
         since we end up here because either seekp or seekg was called
         on the stream using this buffer. That simplifies the code
         in a few places.
      */
      int const failure = off_type(-1);

      if (py_seek.is_none()) {
        throw std::invalid_argument(
          "That Python file object has no 'seek' attribute");
      }

      // we need the read buffer to contain something!
      if (which == std::ios_base::in && !gptr()) {
        if (traits_type::eq_int_type(underflow(), traits_type::eof())) {
          return failure;
        }
      }

      // compute the whence parameter for Python seek
      int whence;
      switch (way) {
        case std::ios_base::beg:
          whence = 0;
          break;
        case std::ios_base::cur:
          whence = 1;
          break;
        case std::ios_base::end:
          whence = 2;
          break;
        default:
          return failure;
      }

      // Let's have a go
      auto result = seekoff_without_calling_python(off, way, which);
      if (!result.second) {
        // we need to call Python
        if (which == std::ios_base::out) overflow();
        if (way == std::ios_base::cur) {
          if      (which == std::ios_base::in)  off -= egptr() - gptr();
          else if (which == std::ios_base::out) off += pptr() - pbase();
        }
        py_seek(off, whence);
        result.first = py_tell().cast<pybind11::int_>();
        if (which == std::ios_base::in) underflow();
      }
      return result.first;
    }

    /// C.f. C++ standard section 27.5.2.4.2
    virtual
    pos_type seekpos(pos_type sp,
                     std::ios_base::openmode which=  std::ios_base::in
                                                   | std::ios_base::out)
    {
      return python_streambuf::seekoff(sp, std::ios_base::beg, which);
    }

  private:
    pybind11::function py_read;
    pybind11::function py_write;
    pybind11::function py_seek;
    pybind11::function py_tell;

    std::size_t buffer_size;

    /* This is actually a Python string and the actual read buffer is
       its internal data, i.e. an array of characters. We use a Boost.Python
       object so as to hold on it: as a result, the actual buffer can't
       go away.
    */
    pybind11::bytes read_buffer;

    /* A mere array of char's allocated on the heap at construction time and
       de-allocated only at destruction time.
    */
    char *write_buffer = nullptr;

    off_type pos_of_read_buffer_end_in_py_file,
             pos_of_write_buffer_end_in_py_file;

    // the farthest place the buffer has been written into
    char *farthest_pptr = nullptr;


    std::pair<off_type, bool> seekoff_without_calling_python(
      off_type off,
      std::ios_base::seekdir way,
      std::ios_base::openmode which)
    {
      const auto failure = std::make_pair<off_type, bool>(off_type(), false);

      // Buffer range and current position
      off_type buf_begin, buf_end, buf_cur, upper_bound;
      off_type pos_of_buffer_end_in_py_file;
      if (which == std::ios_base::in) {
        pos_of_buffer_end_in_py_file = pos_of_read_buffer_end_in_py_file;
        buf_begin = reinterpret_cast<std::streamsize>(eback());
        buf_cur = reinterpret_cast<std::streamsize>(gptr());
        buf_end = reinterpret_cast<std::streamsize>(egptr());
        upper_bound = buf_end;
      }
      else if (which == std::ios_base::out) {
        pos_of_buffer_end_in_py_file = pos_of_write_buffer_end_in_py_file;
        buf_begin = reinterpret_cast<std::streamsize>(pbase());
        buf_cur = reinterpret_cast<std::streamsize>(pptr());
        buf_end = reinterpret_cast<std::streamsize>(epptr());
        farthest_pptr = std::max(farthest_pptr, pptr());
        upper_bound = reinterpret_cast<std::streamsize>(farthest_pptr) + 1;
      }
      else {
        throw xlnt::exception("unreachable");
      }

      // Sought position in "buffer coordinate"
      off_type buf_sought;
      if (way == std::ios_base::cur) {
        buf_sought = buf_cur + off;
      }
      else if (way == std::ios_base::beg) {
        buf_sought = buf_end + (off - pos_of_buffer_end_in_py_file);
      }
      else if (way == std::ios_base::end) {
        return failure;
      }
      else {
        throw xlnt::exception("unreachable");
      }

      // if the sought position is not in the buffer, give up
      if (buf_sought < buf_begin || buf_sought >= upper_bound) return failure;

      // we are in wonderland
      if      (which == std::ios_base::in)  gbump(static_cast<int>(buf_sought - buf_cur));
      else if (which == std::ios_base::out) pbump(static_cast<int>(buf_sought - buf_cur));
      return std::make_pair<off_type, bool>(pos_of_buffer_end_in_py_file + (buf_sought - buf_end), true);
    }
};

std::size_t python_streambuf::default_buffer_size = 1024;

} // namespace xlnt
