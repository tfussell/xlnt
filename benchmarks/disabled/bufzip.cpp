#include <xlnt/xlnt.hpp>
#include <xlnt/serialization/xml_document.hpp>
#include <xlnt/serialization/xml_node.hpp>
#include <xlnt/serialization/xml_serializer.hpp>

void standard()
{
    xlnt::xml_document doc;

    for (int i = 0; i < 1000000; i++) 
    {
	doc.add_child("test");
    }

    xlnt::zip_file archive;
    archive.writestr("sheet.xml", doc.to_string());
}

int main()
{
    standard();
    return 0;
}
