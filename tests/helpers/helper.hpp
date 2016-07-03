#pragma once

#include <pugixml.hpp>
#include <sstream>

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
    
    static comparison_result compare_xml(const pugi::xml_document &expected, const pugi::xml_document &observed)
    {
        return compare_xml(expected.root(), observed.root());
    }

    static comparison_result compare_xml(const std::string &expected, const pugi::xml_document &observed)
    {
        std::string expected_contents = expected;
        
        if(PathHelper::FileExists(expected))
        {
            std::ifstream f(expected);
            std::ostringstream s;
            f >> s.rdbuf();
            
            expected_contents = s.str();
        }
        
		pugi::xml_document expected_xml;
		expected_xml.load(expected_contents.c_str());

        std::ostringstream ss;
        observed.save(ss);
		auto observed_string = ss.str();

        return compare_xml(expected_xml.root(), observed.root());
    }
    
    static comparison_result compare_xml(const std::string &left_contents, const std::string &right_contents)
    {
		pugi::xml_document left_xml;
		left_xml.load(left_contents.c_str());

		pugi::xml_document right_xml;
		right_xml.load(right_contents.c_str());

        return compare_xml(left_xml.root(), right_xml.root());
    }
    
    static comparison_result compare_xml(const pugi::xml_node &left, const pugi::xml_node &right)
    {
        std::string left_temp = left.name();
        std::string right_temp = right.name();
        
        if(left_temp != right_temp)
        {
            return {difference_type::names_differ, left_temp, right_temp};
        }
        
        for(auto &left_attribute : left.attributes())
        {
            left_temp = left_attribute.name();

            if(!right.attribute(left_attribute.name()))
            {
                return {difference_type::missing_attribute, left_temp, "((empty))"};
            }
            
            left_temp = left_attribute.name();
            right_temp = right.attribute(left_attribute.name()).name();
            
            if(left_temp != right_temp)
            {
                return {difference_type::attribute_values_differ, left_temp, right_temp};
            }
        }
        
        if(left.text())
        {
            left_temp = left.text().get();
            
            if(!right.text())
            {
                return {difference_type::missing_text, left_temp, "((empty))"};
            }
            
            right_temp = right.text().get();
            
            if(left_temp != right_temp)
            {
                return {difference_type::text_values_differ, left_temp, right_temp};
            }
        }
        else if(right.text())
        {
            right_temp = right.text().get();
            return {difference_type::text_values_differ, "((empty))", right_temp};
        }
        
        auto right_children = right.children();
        auto right_child_iter = right_children.begin();
        
        for(auto left_child : left.children())
        {
            left_temp = left_child.name();
            
            if(right_child_iter == right_children.end())
            {
                return {difference_type::child_order_differs, left_temp, "((end))"};
            }
            
            auto right_child = *right_child_iter;
            right_child_iter++;
            
            auto child_comparison_result = compare_xml(left_child, right_child);
            
            if(!child_comparison_result)
            {
                return child_comparison_result;
            }
        }
        
        if(right_child_iter != right_children.end())
        {
            right_temp = right_child_iter->name();
            return {difference_type::child_order_differs, "((end))", right_temp};
        }
        
        return {difference_type::equivalent, "", ""};
    }
};
