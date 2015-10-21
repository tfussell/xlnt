#include <xlnt/styles/protection.hpp>

namespace xlnt {

protection::protection() : protection(type::unprotected)
{
    
}

protection::protection(type t) : locked_(t), hidden_(type::inherit)
{
    
}

} // namespace xlnt
