#pragma once

#include <vector>

#include <xlnt/utils/string.hpp>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

class relationship;
class zip_file;

/// <summary>
/// Reads and writes collections of relationshps for a particular file.
/// </summary>
class XLNT_CLASS relationship_serializer
{
public:
    /// <summary>
    /// Construct a serializer which operates on archive.
    /// </summary>
    relationship_serializer(zip_file &archive);
    
    /// <summary>
    /// Return a vector of relationships corresponding to target.
    /// </summary>
    std::vector<relationship> read_relationships(const string &target);
    
    /// <summary>
    /// Write relationships to archive for the given target.
    /// </summary>
    bool write_relationships(const std::vector<relationship> &relationships, const string &target);
    
private:
    /// <summary>
    /// Internal archive which is used for reading and writing.
    /// </summary>
    zip_file &archive_;
};

} // namespace xlnt
