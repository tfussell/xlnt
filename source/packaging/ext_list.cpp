#include <xlnt/packaging/ext_list.hpp>
#include <algorithm>

#include <detail/external/include_libstudxml.hpp>

namespace {
    // send elements straight from parser to serialiser without modification
    // runs until the end of the current open element
    void roundtrip(xml::parser& p, xml::serializer &s)
    {
        int nest_level = 0;
        while (nest_level > 0 || (p.peek() != xml::parser::event_type::end_element && p.peek() != xml::parser::event_type::eof))
        {
            switch (p.next())
            {
            case xml::parser::start_element:
            {
                ++nest_level;
                s.start_element(p.qname());
                for (auto &ele : p.attribute_map())
                {
                    s.attribute(ele.first, ele.second.value);
                    ele.second.handled = true;
                }
                break;
            }
            case xml::parser::end_element:
            {
                --nest_level;
                s.end_element();
                break;
            }
            case xml::parser::start_namespace_decl:
            {
                s.namespace_decl(p.namespace_(), p.prefix());
                break;
            }
            case xml::parser::end_namespace_decl:
            { // nothing required here
                break;
            }
            case xml::parser::characters:
            {
                s.characters(p.value());
                break;
            }
            }
        }
    }
}

namespace xlnt {

ext_list::ext::ext(xml::parser &parser)
{
    // because libstudxml is a streaming parser, this is parsed once to acquire the complete string form, and then a second time for the URI
    std::ostringstream serialisation_stream;
    xml::serializer s(serialisation_stream, "", 0);
    roundtrip(parser, s);
    serialised_value_ = serialisation_stream.str();
    const std::string uri_specifier = R"(uri=")";
    int offset = serialised_value_.find(uri_specifier) + uri_specifier.size();
    int len = serialised_value_.find('"', offset) - offset;
    extension_ID_ = uri(serialised_value_.substr(offset, len));
}

ext_list::ext::ext(const uri &ID, const std::string &serialised)
    : extension_ID_(ID), serialised_value_(serialised)
{}

void ext_list::ext::serialise(xml::serializer &serialiser)
{
    std::istringstream ser(serialised_value_);
    xml::parser p(ser, "", xml::parser::receive_default);
    int nest_level = 0;
    roundtrip(p, serialiser);
}

ext_list::ext_list(xml::parser &parser)
{
    // begin with the start element already parsed
    while (parser.peek() == xml::parser::start_element)
    {
        extensions_.push_back(ext(parser));
    }
    // end without parsing the end element
}

void ext_list::serialize(xml::serializer &serialiser, std::string namesp)
{
    serialiser.start_element("extLst", namesp);
    for (auto &ext : extensions_)
    {
        ext.serialise(serialiser);
    }
    serialiser.end_element();
}

void ext_list::add_extension(const uri &ID, const std::string &element)
{
    extensions_.push_back(ext{ID, element});
}

bool ext_list::has_extension(const uri &extension_uri) const
{
    return extensions_.end() != std::find_if(extensions_.begin(), extensions_.end(), 
                                             [&extension_uri](const ext &ext) { return extension_uri == ext.extension_ID_; });
}

const ext_list::ext &ext_list::extension(const uri &extension_uri) const
{
    return *std::find_if(extensions_.begin(), extensions_.end(),
        [&extension_uri](const ext &ext) { return extension_uri == ext.extension_ID_; });
}

const std::vector<ext_list::ext> &ext_list::extensions() const
{
    return extensions_;
}

}