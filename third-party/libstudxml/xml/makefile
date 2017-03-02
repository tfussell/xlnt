# file      : xml/makefile
# copyright : Copyright (c) 2013-2017 Code Synthesis Tools CC
# license   : MIT; see accompanying LICENSE file

include $(dir $(lastword $(MAKEFILE_LIST)))../build/bootstrap.make

sources := qname.cxx parser.cxx serializer.cxx value-traits.cxx
xml_headers := content exception forward qname parser serializer value-traits \
version

# Expat.
#
expat_sources := xmlparse.c xmlrole.c xmltok.c
expat_headers := expat.h expat_external.h 
expat_dist    := ascii.h asciitab.h config.h iasciitab.h internal.h \
latin1tab.h nametab.h utf8tab.h xmlrole.h xmltok.h xmltok_impl.c \
xmltok_impl.h xmltok_ns.c README LICENSE

expat_sources := $(addprefix details/expat/,$(expat_sources))
expat_headers := $(addprefix details/expat/,$(expat_headers))
expat_dist    := $(addprefix details/expat/,$(expat_dist))

# Genx.
#
genx_sources := details/genx/genx.c details/genx/char-props.c
genx_headers := details/genx/genx.h
genx_dist    := details/genx/README details/genx/LICENSE

c_tun     := $(expat_sources) $(genx_sources)
c_obj     := $(addprefix $(out_base)/,$(c_tun:.c=.o))
c_od      := $(c_obj:.o=.o.d)

cxx_tun   := $(sources)
cxx_obj   := $(addprefix $(out_base)/,$(cxx_tun:.cxx=.o))
cxx_od    := $(cxx_obj:.o=.o.d)

studxml.l             := $(out_base)/studxml.l
studxml.l.cpp-options := $(out_base)/studxml.l.cpp-options

default   := $(out_base)/
dist      := $(out_base)/.dist
install   := $(out_base)/.install
clean     := $(out_base)/.clean

# Build.
#
$(studxml.l): $(c_obj) $(cxx_obj)

$(c_obj) $(cxx_obj) $(c_od) $(cxx_od): $(studxml.l.cpp-options) \
  $(out_base)/details/config.h
$(studxml.l.cpp-options): value := -I$(out_root) -I$(src_root)

$(call include,$(bld_root)/ld/configuration-lib.make) # ld_lib_type

$(out_base)/details/config.h: | $(out_base)/details/.
	@echo '/* file      : xml/details/config.h'                >$@
	@echo ' * note      : automatically generated'            >>$@
	@echo ' */'                                               >>$@
	@echo '#ifndef XML_DETAILS_CONFIG_H'                      >>$@
	@echo '#define XML_DETAILS_CONFIG_H'                      >>$@
	@echo ''                                                  >>$@
	@echo '#define LIBSTUDXML_BYTEORDER 1234'                 >>$@
ifeq ($(ld_lib_type),archive)
	@echo '#define LIBSTUDXML_STATIC_LIB'                     >>$@
endif
	@echo ''                                                  >>$@
	@echo '#endif /* XML_DETAILS_CONFIG_H */'                 >>$@

$(call include-dep,$(c_od) $(cxx_od),$(c_obj) $(cxx_obj),$(out_base)/details/config.h)

# Convenience alias for default target.
#
$(out_base)/: $(studxml.l)

# Dist.
#

# Set it for the out_root so that it is visible in examples/tests. See
# Makefile.am for why we need it there.
#
$(our_root)/%.dist: export xml_headers := $(xml_headers) 

$(dist): export sources := $(sources)
$(dist): export expat_sources := $(expat_sources)
$(dist): export expat_headers := $(expat_headers)
$(dist): export genx_sources := $(genx_sources)
$(dist): export genx_headers := $(genx_headers)
$(dist): export headers := $(subst $(src_base)/,,$(shell find $(src_base) \
-name '*.hxx' -o -name '*.ixx' -o -name '*.txx'))
$(dist): data_dist := details/config-vc.h $(expat_dist) $(genx_dist)
$(dist): export extra_dist := $(data_dist) libstudxml-vc9.vcproj \
libstudxml-vc10.vcxproj libstudxml-vc10.vcxproj.filters \
libstudxml-vc11.vcxproj libstudxml-vc11.vcxproj.filters \
libstudxml-vc12.vcxproj libstudxml-vc12.vcxproj.filters
$(dist): export interface_version = $(shell sed -e \
's/^\([0-9]*\.[0-9]*\).*/\1/' $(src_root)/version)

$(dist):
	$(call dist-data,$(sources) $(expat_sources) $(genx_sources))
	$(call dist-data,$(headers) $(xml_headers) $(expat_headers) $(genx_headers))
	$(call dist-data,$(data_dist) details/config.h.in)
	$(call meta-vc9proj,libstudxml-vc9.vcproj)
	$(call meta-vc10proj,libstudxml-vc10.vcxproj)
	$(call meta-vc10proj,libstudxml-vc11.vcxproj)
	$(call meta-vc12proj,libstudxml-vc12.vcxproj)
	$(call meta-automake)

# Clean.
#
$(clean): $(studxml.l).o.clean       \
  $(studxml.l.cpp-options).clean     \
  $(addsuffix .cxx.clean,$(c_obj))   \
  $(addsuffix .cxx.clean,$(c_od))    \
  $(addsuffix .cxx.clean,$(cxx_obj)) \
  $(addsuffix .cxx.clean,$(cxx_od))
	$(call message,rm $$1,rm -f $$1,$(out_base)/details/config.h)

# Generated .gitignore.
#
ifeq ($(out_base),$(src_base))
$(studxml.l): | $(out_base)/.gitignore

$(out_base)/.gitignore: files := details/config.h
$(clean): $(out_base)/.gitignore.clean

$(call include,$(bld_root)/git/gitignore.make)
endif

# How to.
#
$(call include,$(bld_root)/dist.make)
$(call include,$(bld_root)/install.make)
$(call include,$(bld_root)/meta/vc9proj.make)
$(call include,$(bld_root)/meta/vc10proj.make)
$(call include,$(bld_root)/meta/vc11proj.make)
$(call include,$(bld_root)/meta/vc12proj.make)
$(call include,$(bld_root)/meta/automake.make)

$(call include,$(bld_root)/c/c-o.make)
$(call include,$(bld_root)/c/c-d.make)
$(call include,$(bld_root)/cxx/cxx-o.make)
$(call include,$(bld_root)/cxx/cxx-d.make)
$(call include,$(bld_root)/cxx/o-l.make)
