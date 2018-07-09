#include <xlnt/packaging/ext_list.hpp>
#include <algorithm>

#include <detail/external/include_libstudxml.hpp>

namespace {
    // send elements straight from parser to serialiser without modification
    // runs until the end of the current open element
    xlnt::uri roundtrip(xml::parser& p, xml::serializer &s)
    {
        xlnt::uri ext_uri;
        int nest_level = 0;
        while (nest_level > 0 || (p.peek() != xml::parser::event_type::end_element && p.peek() != xml::parser::event_type::eof))
        {
            switch (p.next())
            {
            case xml::parser::start_element:
            {
                ++nest_level;
                auto attribs = p.attribute_map();
                s.start_element(p.qname());
                if (nest_level == 1)
                {
                    ext_uri = xlnt::uri(attribs.at(xml::qname("uri")).value);
                }
                auto current_ns = p.namespace_();
                p.peek(); // to look into the new namespace
                auto new_ns = p.namespace_(); // only before attributes?
                if (new_ns != current_ns)
                {
                    auto pref = p.prefix();
                    s.namespace_decl(new_ns, pref);
                }
                for (auto &ele : attribs)
                {
                    s.attribute(ele.first.string(), ele.second.value);
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
            case xml::parser::eof:
                return ext_uri;
            case xml::parser::start_attribute:
            case xml::parser::end_attribute:
            default:
                break;
            }
        }
        return ext_uri;
    }
}

namespace xlnt {

ext_list::ext::ext(xml::parser &parser, const std::string& ns)
{
    std::ostringstream serialisation_stream;
    xml::serializer s(serialisation_stream, "", 0);
    s.start_element(xml::qname(ns, "wrap")); // wrapper for the xmlns declaration
    s.namespace_decl(ns, "");
    extension_ID_ = roundtrip(parser, s);
    s.end_element(xml::qname(ns, "wrap"));
    serialised_value_ = serialisation_stream.str();
}

ext_list::ext::ext(const uri &ID, const std::string &serialised)
    : extension_ID_(ID), serialised_value_(serialised)
{}

void ext_list::ext::serialise(xml::serializer &serialiser, const std::string& ns)
{
    std::istringstream ser(serialised_value_);
    xml::parser p(ser, "", xml::parser::receive_default);
    p.next_expect(xml::parser::event_type::start_element, xml::qname(ns, "wrap"));
    roundtrip(p, serialiser);
    p.next_expect(xml::parser::event_type::end_element, xml::qname(ns, "wrap"));
}

ext_list::ext_list(xml::parser &parser, const std::string& ns)
{
    // begin with the start element already parsed
    while (parser.peek() == xml::parser::start_element)
    {
        extensions_.push_back(ext(parser, ns));
    }
    // end without parsing the end element
}

void ext_list::serialize(xml::serializer &serialiser, const std::string& ns)
{
    serialiser.start_element(ns, "extLst");
    for (auto &ext : extensions_)
    {
        ext.serialise(serialiser, ns);
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

bool ext_list::operator==(const ext_list &rhs) const
{
    return extensions_ == rhs.extensions_;
}

}