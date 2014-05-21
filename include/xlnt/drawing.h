#pragma once

namespace xlnt {

class worksheet;
struct drawing_struct;
    
class drawing
{
public:
    drawing();
    
private:
    friend class worksheet;
    drawing(drawing_struct *root);
    drawing_struct *root_;
};

} // namespace xlnt
