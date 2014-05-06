#include <sstream>

#include "uri.h"

namespace xlnt {

namespace {
#ifdef _WIN32
const char platform_path_separator = '\\';
#else
const char platform_path_separator = '/';
#endif

int get_DefaultPort(const std::string &scheme)
{
	if(scheme == "ftp") return 21;
	if(scheme == "sftp") return 22;
	if(scheme == "ssh") return 22;
	if(scheme == "telnet") return 23;
	if(scheme == "smtp") return 25;
	if(scheme == "dns") return 53;
	if(scheme == "gopher") return 70;
	if(scheme == "http") return 80;
	if(scheme == "pop3") return 110;
	if(scheme == "nntp") return 119;
	if(scheme == "snmp") return 161;
	if(scheme == "imap") return 143;
	if(scheme == "irc") return 194;
	if(scheme == "https") return 443;
	if(scheme == "smtps") return 465;
	return -1;
}
}

uri::uri(const std::string &uri_string, uri_kind uri_kind)
	: original_string_(uri_string),
	  kind_(uri_kind)
{
	Parse();
}

uri &uri::operator=(const uri &rhs)
{
    original_string_ = rhs.original_string_;
    kind_ = rhs.kind_;
    Parse();
    return *this;
}

bool uri::IsAbsoluteuri() const
{
	if(kind_ == uri_kind::Absolute)
	{
		return true;
	}

	if(kind_ == uri_kind::Relative)
	{
		return false;
	}

	if(original_string_.size() == 0)
	{
		return false;
	}

	auto scheme_separator = std::find(original_string_.begin(), original_string_.end(), ':');
	if(scheme_separator == original_string_.end())
	{
		return false;
	}

	if(original_string_[0] == '.')
	{
		return false;
	}

	return true;
}

void uri::Parse()
{
	if(IsAbsoluteuri())
	{
		ParseAbsolute();
	}
	else
	{
		ParseRelative();
	}
}

void uri::ParseAbsolute()
{
	auto scheme_separator = std::find(original_string_.begin(), original_string_.end(), ':');

	if(scheme_separator == original_string_.end())
	{
		throw std::runtime_error("absolute uri must contain a scheme followed by a colon");
	}

	scheme_ = std::string(original_string_.begin(), scheme_separator);

	port_ = get_DefaultPort(scheme_);

	std::string hier_part(scheme_separator + 1, original_string_.end());

	auto query_separator = std::find(hier_part.begin(), hier_part.end(), '?');
	if(query_separator != hier_part.end())
	{
		query_ = std::string(query_separator + 1, hier_part.end());		
		hier_part = std::string(hier_part.begin(), query_separator);
	}

	auto fragment_separator = std::find(hier_part.begin(), hier_part.end(), '#');
	if(fragment_separator != hier_part.end())
	{
		fragment_ = std::string(fragment_separator + 1, hier_part.end());		
		hier_part = std::string(hier_part.begin(), fragment_separator);
	}
	else
	{
		fragment_separator = std::find(query_.begin(), query_.end(), '#');
		if(fragment_separator != query_.end())
		{
			fragment_ = std::string(fragment_separator + 1, query_.end());			
			query_ = std::string(query_.begin(), fragment_separator);
		}
	}

	if(hier_part.size() > 2 && hier_part[0] == '/' && hier_part[1] == '/')
	{
		auto path_separator = std::find(hier_part.begin() + 2, hier_part.end(), '/');
		if(path_separator != hier_part.end())
		{
			authority_ = std::string(hier_part.begin() + 2, path_separator);
			absolute_path_ = std::string(path_separator, hier_part.end());
			if(absolute_path_[0] == '/')
			{
				if(absolute_path_[1] == '/')
				{
					throw std::runtime_error("path can't start with //");
				}
			}
		}
		else
		{
			authority_ = hier_part.substr(2);
			absolute_path_ = "";
		}

		ParseAuthority();

		is_file_ = scheme_ == "file";

		char abs_path_separator = absolute_path_.find('/') != std::string::npos ? '/' : '\\';

		if(is_file_)
		{
			local_path_ = std::string(2, platform_path_separator) + authority_;
		}

		if(absolute_path_.size() > 0)
		{
			std::stringstream ss(absolute_path_);
			std::string part;
			while(std::getline(ss, part, abs_path_separator))
			{
				if(segments_.size() > 0)
				{
					if(is_file_)
					{
						segments_.back().append(1, platform_path_separator);
						local_path_.append(1, platform_path_separator);
					}
					else
					{
						segments_.back().append(1, abs_path_separator);
						local_path_.append(1, abs_path_separator);
					}
				}
				segments_.push_back(part);
				local_path_.append(part);
			}
		}

		user_escaped_ = false;
		is_unc_ = is_file_ && host_.size() > 0;
		is_default_port_ = get_DefaultPort(scheme_) == port_;
		is_loopback_ = host_ == "127.0.0.1" || host_ == "localhost" || host_ == "loopback" || host_.size() == 0;
		path_and_query_ = absolute_path_ + "?" + query_;
		absolute_uri_ = scheme_ + "://";
		if(user_info_.size() > 0)
		{
			absolute_uri_ += user_info_ + "@";
		}
		absolute_uri_ += host_;
		if(!is_default_port_)
		{
			absolute_uri_ += ":" + std::to_string(port_);
		}
		absolute_uri_ += absolute_path_;
		if(query_.size() > 0)
		{
			absolute_uri_ += "?" + query_;
		}
		if(fragment_.size() > 0)
		{
			absolute_uri_ += "#" + fragment_;
		}
	}
}

void uri::ParseRelative()
{
	absolute_path_ = "";
	absolute_uri_ = "";
	authority_ = "";
	dns_safe_host_ = "";
	fragment_ = "";
	host_ = "";
	host_name_type_ = uri_host_name_type::Unknown;
	is_default_port_ = false;
	is_file_ = false;
	is_loopback_ = false;
	is_unc_ = false;
	local_path_ = "";
	path_and_query_ = "";
	port_ = -1;
	query_ = "";
	scheme_ = "";
	user_escaped_ = false;
	user_info_ = "";
}

void uri::ParseAuthority()
{
	if(authority_.size() == 0)
	{
		return;
	}

	host_ = authority_;

	auto user_info_separator = std::find(host_.begin(), host_.end(), '@');
	if(user_info_separator != host_.end())
	{
		user_info_ = std::string(host_.begin(), user_info_separator);
		host_ = std::string(user_info_separator + 1, host_.end());
	}

	auto port_separator_index = host_.find_last_of(':');
	if(port_separator_index != std::string::npos)
	{
		port_ = std::stoi(host_.substr(port_separator_index + 1));
		host_ = host_.substr(0, port_separator_index);
	}

	dns_safe_host_ = host_;

	if(host_.find(':') == std::string::npos)
	{
		host_name_type_ = uri_host_name_type::Dns;
	}
	else
	{
		host_name_type_ = uri_host_name_type::IPv4;
	}
}

std::string uri::get_AbsolutePath() const
{
	ThrowIfNotAbsolute();
	return absolute_path_;
}

std::string uri::get_Absoluteuri() const
{
	ThrowIfNotAbsolute();
	return absolute_uri_;
}

std::string uri::get_Authority() const
{
	ThrowIfNotAbsolute();
	return authority_;
}

std::string uri::get_DnsSafeHost() const
{
	ThrowIfNotAbsolute();
	return dns_safe_host_;
}

std::string uri::get_Fragment() const
{
	ThrowIfNotAbsolute();
	return fragment_;
}

std::string uri::get_Host() const
{
	ThrowIfNotAbsolute();
	return host_;
}

uri_host_name_type uri::get_HostNameType() const
{
	ThrowIfNotAbsolute();
	return host_name_type_;
}

bool uri::IsDefaultPort() const
{
	ThrowIfNotAbsolute();
	return is_default_port_;
}

bool uri::IsFile() const
{
	ThrowIfNotAbsolute();
	return is_file_;
}

bool uri::IsLoopback() const
{
	ThrowIfNotAbsolute();
	return is_loopback_;
}

bool uri::IsUnc() const
{
	ThrowIfNotAbsolute();
	return is_unc_;
}

std::string uri::get_LocalPath() const
{
	ThrowIfNotAbsolute();
	return local_path_;
}

std::string uri::get_OriginalString() const
{
	return original_string_;
}

std::string uri::get_PathAndQuery() const
{
	ThrowIfNotAbsolute();
	return path_and_query_;
}

int uri::get_Port() const
{
	ThrowIfNotAbsolute();
	return port_;
}

std::string uri::get_Query() const
{
	ThrowIfNotAbsolute();
	return query_;
}

std::string uri::get_Scheme() const
{
	ThrowIfNotAbsolute();
	return scheme_;
}

std::vector<std::string> uri::get_Segments() const
{
	ThrowIfNotAbsolute();
	return segments_;
}

bool uri::get_UserEscaped() const
{
	return user_escaped_;
}

std::string uri::get_UserInfo() const
{
	ThrowIfNotAbsolute();
	return user_info_;
}

void uri::ThrowIfNotAbsolute() const
{
	if(!IsAbsoluteuri())
	{
		throw std::runtime_error("InvalidOperationException: This instance represents a relative URI, and this property is valid only for absolute URIs.");
	}
}

} // namespace xlnt
