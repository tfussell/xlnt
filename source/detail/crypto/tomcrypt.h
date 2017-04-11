#ifndef TOMCRYPT_H_
#define TOMCRYPT_H_
#include <assert.h>
#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <stddef.h>
#include <time.h>
#include <ctype.h>
#include <limits.h>

/* use configuration data */
#include "tomcrypt_custom.h"

#ifdef __cplusplus
extern "C" {
#endif

/* version */
#define CRYPT   0x0117
#define SCRYPT  "1.17"

/* max size of either a cipher/hash block or symmetric key [largest of the two] */
#define MAXBLOCKSIZE  128

/* descriptor table size */
#define TAB_SIZE      32

#include "tomcrypt_cfg.h"
#include "tomcrypt_macros.h"
#include "tomcrypt_cipher.h"
#include "tomcrypt_math.h"
#include "tomcrypt_misc.h"

#ifdef __cplusplus
   }
#endif

#endif /* TOMCRYPT_H_ */


/* $Source$ */
/* $Revision$ */
/* $Date$ */
