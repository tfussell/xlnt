#pragma once

#include <string>

#include "xlnt_config.hpp"

namespace xlnt {

class manifest;
class xml_document;

class XLNT_CLASS manifest_serializer
{
  public:
    manifest_serializer(manifest &m);

    void read_manifest(const xml_document &xml);
    xml_document write_manifest() const;

    string determine_document_type() const;

  private:
    manifest &manifest_;
};

}; // namespace xlnt
