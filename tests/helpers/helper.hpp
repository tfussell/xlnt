#pragma once

#include <sstream>
#include <pugixml.hpp>

#include "path_helper.hpp"

class Helper
{
public:
    static bool EqualsFileContent(const std::string &reference_file, const std::string &fixture)
    {
        std::string fixture_content;

        if(false)//PathHelper::FileExists(fixture))
        {
            std::fstream fixture_file;
            fixture_file.open(fixture);
            std::stringstream ss;
            ss << fixture_file.rdbuf();
            fixture_content = ss.str();
        }
        else
        {
            fixture_content = fixture;
        }

        std::string expected_content;

        if(PathHelper::FileExists(reference_file))
        {
            std::fstream file;
            file.open(reference_file);
            std::stringstream ss;
            ss << file.rdbuf();
            expected_content = ss.str();
        }
        else
        {
            throw std::runtime_error("file not found");
        }

      {
          pugi::xml_document doc;
          doc.load(fixture_content.c_str());
          std::stringstream ss;
          doc.save(ss);
          fixture_content = ss.str();
      }

      {
          pugi::xml_document doc;
          doc.load(expected_content.c_str());
          std::stringstream ss;
          doc.save(ss);
          expected_content = ss.str();
      }

        return expected_content == fixture_content;
    }
};
