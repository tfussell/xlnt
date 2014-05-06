#include "part.h"

namespace xlnt {

class part_impl
{
public:
    /// <summary>
    /// Initializes a new instance of the part class with a specified parent Package, part URI, MIME content type, and compression_option.
    /// </summary>
    part_impl(package &package, const uri &uri_part, const std::string &mime_type = "", compression_option compression = compression_option::NotCompressed)
        : package_(package),
        uri_(uri_part),
        content_type_(mime_type),
        compression_option_(compression)
    {}

    part_impl(package &package, const uri &uri, opcContainer *container)
        : package_(package),
        uri_(uri),
        container_(container)
    {}

    /// <summary>
    /// gets the compression option of the part content stream.
    /// </summary>
    compression_option get_compression_option() const { return compression_option_; }

    /// <summary>
    /// gets the MIME type of the content stream.
    /// </summary>
    std::string get_content_type() const;

    /// <summary>
    /// gets the parent Package of the part.
    /// </summary>
    package &get_package() const { return package_; }

    /// <summary>
    /// gets the URI of the part.
    /// </summary>
    uri get_uri() const { return uri_; }

    /// <summary>
    /// Creates a part-level relationship between this part to a specified target part or external resource.
    /// </summary>
    std::shared_ptr<relationship> create_relationship(const uri &target_uri, target_mode target_mode, const std::string &relationship_type);

    /// <summary>
    /// Deletes a specified part-level relationship.
    /// </summary>
    void delete_relationship(const std::string &id);

    /// <summary>
    /// Returns the relationship that has a specified Id.
    /// </summary>
    relationship get_relationship(const std::string &id);

    /// <summary>
    /// Returns a collection of all the relationships that are owned by this part.
    /// </summary>
    relationship_collection get_relationships();

    /// <summary>
    /// Returns a collection of the relationships that match a specified RelationshipType.
    /// </summary>
    relationship_collection get_relationship_by_type(const std::string &relationship_type);

    std::string read()
    {
        std::string ss;
        auto part_stream = opcContainerOpenInputStream(container_, (xmlChar*)get_uri().get_OriginalString().c_str());
        std::array<xmlChar, 1024> buffer;
        auto bytes_read = opcContainerReadInputStream(part_stream, buffer.data(), static_cast<opc_uint32_t>(buffer.size()));
        if(bytes_read > 0)
        {
            ss.append(std::string(buffer.begin(), buffer.begin() + bytes_read));
            while(bytes_read == buffer.size())
            {
                auto bytes_read = opcContainerReadInputStream(part_stream, buffer.data(), static_cast<opc_uint32_t>(buffer.size()));
                ss.append(std::string(buffer.begin(), buffer.begin() + bytes_read));
            }
        }
        opcContainerCloseInputStream(part_stream);
        return ss;
    }

    void write(const std::string &data)
    {
        auto name = get_uri().get_OriginalString();
        auto name_pointer = name.c_str();

        auto part_stream = opcContainerCreateOutputStream(container_, (xmlChar*)name_pointer, opcCompressionOption_t::OPC_COMPRESSIONOPTION_NORMAL);

        std::stringstream ss(data);
        std::array<xmlChar, 1024> buffer;

        while(ss)
        {
            ss.get((char*)buffer.data(), 1024);
            auto count = ss.gcount();
            if(count > 0)
            {
                opcContainerWriteOutputStream(part_stream, buffer.data(), static_cast<opc_uint32_t>(count));
            }
        }
        opcContainerCloseOutputStream(part_stream);
    }

    /// <summary>
    /// Returns a value that indicates whether this part owns a relationship with a specified Id.
    /// </summary>
    bool relationship_exists(const std::string &id) const;

    /// <summary>
    /// Returns true if the given Id string is a valid relationship identifier.
    /// </summary>
    bool is_valid_xml_id(const std::string &id);

    void operator=(const part_impl &other);

    compression_option compression_option_;
    std::string content_type_;
    package &package_;
    uri uri_;
    opcContainer *container_;
};

part::part() : impl_(nullptr)
{

}

part::part(package &package, const uri &uri, opcContainer *container) : impl_(std::make_shared<part_impl>(package, uri, container))
{

}

std::string part::get_content_type() const
{
	return "";
}

std::string part::read()
{
    if(impl_.get() == nullptr)
    {
        return "";
    }

    return impl_->read();
}

void part::write(const std::string &data)
{
    if(impl_.get() == nullptr)
    {
        return;
    }

    return impl_->write(data);
}

bool part::operator==(const part &comparand) const
{
    return impl_.get() == comparand.impl_.get();
}

bool part::operator==(const nullptr_t &) const
{
    return impl_.get() == nullptr;
}

} // namespace xlnt
