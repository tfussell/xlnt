#pragma once

#include <sstream>
#include <pugixml.hpp>

#include "path_helper.hpp"

class Helper
{
public:
    enum class difference_type
    {
        names_differ,
        missing_attribute,
        attribute_values_differ,
        missing_text,
        text_values_differ,
        missing_child,
        child_order_differs,
        equivalent,
    };
    
    struct comparison_result
    {
        difference_type difference;
        std::string value_left;
        std::string value_right;
        operator bool() const { return difference == difference_type::equivalent; }
    };
    
    static comparison_result compare_xml(const std::string &left_contents, const std::string &right_contents)
    {
        pugi::xml_document left_doc;
        left_doc.load(left_contents.c_str());
        
        pugi::xml_document right_doc;
        right_doc.load(right_contents.c_str());
        
        return compare_xml(left_doc.root(), right_doc.root());
    }
    
    static comparison_result compare_xml(const pugi::xml_node &left, const pugi::xml_node &right)
    {
        std::string left_temp = left.name();
        std::string right_temp = right.name();
        
        if(left_temp != right_temp)
        {
            return {difference_type::names_differ, left_temp, right_temp};
        }
        
        for(auto left_attribute : left.attributes())
        {
            left_temp = left_attribute.name();
            auto right_attribute = right.attribute(left_attribute.name());
            
            if(right_attribute == nullptr)
            {
                return {difference_type::missing_attribute, left_temp, "((empty))"};
            }
            
            left_temp = left_attribute.value();
            right_temp = right_attribute.value();
            
            if(left_temp != right_temp)
            {
                return {difference_type::attribute_values_differ, left_temp, right_temp};
            }
        }
        
        if(left.text() != nullptr)
        {
            left_temp = left.text().as_string();
            
            if(right.text() == nullptr)
            {
                return {difference_type::missing_text, left_temp, "((empty))"};
            }
            
            right_temp = right.text().as_string();
            
            if(left_temp != right_temp)
            {
                return {difference_type::text_values_differ, left_temp, right_temp};
            }
        }
        else if(right.text() != nullptr)
        {
            right_temp = right.text().as_string();
            return {difference_type::text_values_differ, "((empty))", right_temp};
        }
        
        auto right_child_iter = right.children().begin();
        for(auto left_child : left.children())
        {
            left_temp = left_child.name();
            
            if(right_child_iter == right.children().end())
            {
                return {difference_type::child_order_differs, left_temp, "((end))"};
            }
            
            auto right_child = *right_child_iter;
            right_child_iter++;
            
            if(left_child.type() == pugi::xml_node_type::node_cdata || left_child.type() == pugi::xml_node_type::node_pcdata)
            {
                continue; // ignore text children, these have already been compared
            }
            
            auto child_comparison_result = compare_xml(left_child, right_child);
            
            if(!child_comparison_result)
            {
                return child_comparison_result;
            }
        }
        
        if(right_child_iter != right.children().end())
        {
            right_temp = right_child_iter->name();
            return {difference_type::child_order_differs, "((end))", right_temp};
        }
        
        return {difference_type::equivalent, "", ""};
    }
    
    static bool EqualsFileContent(const std::string &reference_file, const std::string &observed_content)
    {
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
        
        pugi::xml_document doc_observed;
        doc_observed.load(observed_content.c_str());
        
        pugi::xml_document doc_expected;
        doc_expected.load(expected_content.c_str());
        
        return compare_xml(doc_expected, doc_observed);
    }
};
