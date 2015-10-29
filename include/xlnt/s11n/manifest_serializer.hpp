#pragma once

#include <string>

namespace xlnt {
    
class manifest;
class workbook;
class xml_document;
class zip_file;

class manifest_serializer
{
public:
    manifest_serializer(manifest &m);
    
    bool read_mainfest(const xml_document &xml);
    bool write_manifest(xml_document &xml);
    
private:
    manifest &manifest_;
};
    
}; // namespace xlnt
