#ifdef LTC_RIJNDAEL
struct rijndael_key {
   ulong32 eK[60], dK[60];
   int Nr;
};
#endif

typedef union Symmetric_key {
#ifdef LTC_RIJNDAEL
   struct rijndael_key rijndael;
#endif
   void   *data;
} symmetric_key;

#ifdef LTC_ECB_MODE
/** A block cipher ECB structure */
typedef struct {
   /** The index of the cipher chosen */
   int                 cipher,
   /** The block size of the given cipher */
                       blocklen;
   /** The scheduled key */
   symmetric_key       key;
} symmetric_ECB;
#endif

#ifdef LTC_CFB_MODE
/** A block cipher CFB structure */
typedef struct {
   /** The index of the cipher chosen */
   int                 cipher,
   /** The block size of the given cipher */
                       blocklen,
   /** The padding offset */
                       padlen;
   /** The current IV */
   unsigned char       IV[MAXBLOCKSIZE],
   /** The pad used to encrypt/decrypt */
                       pad[MAXBLOCKSIZE];
   /** The scheduled key */
   symmetric_key       key;
} symmetric_CFB;
#endif

#ifdef LTC_OFB_MODE
/** A block cipher OFB structure */
typedef struct {
   /** The index of the cipher chosen */
   int                 cipher,
   /** The block size of the given cipher */
                       blocklen,
   /** The padding offset */
                       padlen;
   /** The current IV */
   unsigned char       IV[MAXBLOCKSIZE];
   /** The scheduled key */
   symmetric_key       key;
} symmetric_OFB;
#endif

#ifdef LTC_CBC_MODE
/** A block cipher CBC structure */
typedef struct {
   /** The index of the cipher chosen */
   int                 cipher,
   /** The block size of the given cipher */
                       blocklen;
   /** The current IV */
   unsigned char       IV[MAXBLOCKSIZE];
   /** The scheduled key */
   symmetric_key       key;
} symmetric_CBC;
#endif


#ifdef LTC_CTR_MODE
/** A block cipher CTR structure */
typedef struct {
   /** The index of the cipher chosen */
   int                 cipher,
   /** The block size of the given cipher */
                       blocklen,
   /** The padding offset */
                       padlen,
   /** The mode (endianess) of the CTR, 0==little, 1==big */
                       mode,
   /** counter width */
                       ctrlen;

   /** The counter */
   unsigned char       ctr[MAXBLOCKSIZE],
   /** The pad used to encrypt/decrypt */
                       pad[MAXBLOCKSIZE];
   /** The scheduled key */
   symmetric_key       key;
} symmetric_CTR;
#endif


#ifdef LTC_LRW_MODE
/** A LRW structure */
typedef struct {
    /** The index of the cipher chosen (must be a 128-bit block cipher) */
    int               cipher;

    /** The current IV */
    unsigned char     IV[16],

    /** the tweak key */
                      tweak[16],

    /** The current pad, it's the product of the first 15 bytes against the tweak key */
                      pad[16];

    /** The scheduled symmetric key */
    symmetric_key     key;

#ifdef LTC_LRW_TABLES
    /** The pre-computed multiplication table */
    unsigned char     PC[16][256][16];
#endif
} symmetric_LRW;
#endif

#ifdef LTC_F8_MODE
/** A block cipher F8 structure */
typedef struct {
   /** The index of the cipher chosen */
   int                 cipher,
   /** The block size of the given cipher */
                       blocklen,
   /** The padding offset */
                       padlen;
   /** The current IV */
   unsigned char       IV[MAXBLOCKSIZE],
                       MIV[MAXBLOCKSIZE];
   /** Current block count */
   ulong32             blockcnt;
   /** The scheduled key */
   symmetric_key       key;
} symmetric_F8;
#endif


/** cipher descriptor table, last entry has "name == NULL" to mark the end of table */
extern struct ltc_cipher_descriptor {
   /** name of cipher */
   char *name;
   /** internal ID */
   unsigned char ID;
   /** min keysize (octets) */
   int  min_key_length,
   /** max keysize (octets) */
        max_key_length,
   /** block size (octets) */
        block_length,
   /** default number of rounds */
        default_rounds;
   /** Setup the cipher
      @param key         The input symmetric key
      @param keylen      The length of the input key (octets)
      @param num_rounds  The requested number of rounds (0==default)
      @param skey        [out] The destination of the scheduled key
      @return CRYPT_OK if successful
   */
   void  (*setup)(const unsigned char *key, int keylen, int num_rounds, symmetric_key *skey);
   /** Encrypt a block
      @param pt      The plaintext
      @param ct      [out] The ciphertext
      @param skey    The scheduled key
      @return CRYPT_OK if successful
   */
   void (*ecb_encrypt)(const unsigned char *pt, unsigned char *ct, symmetric_key *skey);
   /** Decrypt a block
      @param ct      The ciphertext
      @param pt      [out] The plaintext
      @param skey    The scheduled key
      @return CRYPT_OK if successful
   */
   void (*ecb_decrypt)(const unsigned char *ct, unsigned char *pt, symmetric_key *skey);

   /** Terminate the context
      @param skey    The scheduled key
   */
   void (*done)(symmetric_key *skey);

   /** Determine a key size
       @param keysize    [in/out] The size of the key desired and the suggested size
       @return CRYPT_OK if successful
   */
   void  (*keysize)(int *keysize);
} cipher_descriptor[];

#ifdef LTC_ECB_MODE
void ecb_start(int cipher, const unsigned char *key,
              int keylen, int num_rounds, symmetric_ECB *ecb);
void ecb_encrypt(const unsigned char *pt, unsigned char *ct, unsigned long len, symmetric_ECB *ecb);
void ecb_decrypt(const unsigned char *ct, unsigned char *pt, unsigned long len, symmetric_ECB *ecb);
void ecb_done(symmetric_ECB *ecb);
#endif

#ifdef LTC_CFB_MODE
int cfb_start(int cipher, const unsigned char *IV, const unsigned char *key,
              int keylen, int num_rounds, symmetric_CFB *cfb);
int cfb_encrypt(const unsigned char *pt, unsigned char *ct, unsigned long len, symmetric_CFB *cfb);
int cfb_decrypt(const unsigned char *ct, unsigned char *pt, unsigned long len, symmetric_CFB *cfb);
int cfb_getiv(unsigned char *IV, unsigned long *len, symmetric_CFB *cfb);
int cfb_setiv(const unsigned char *IV, unsigned long len, symmetric_CFB *cfb);
int cfb_done(symmetric_CFB *cfb);
#endif

#ifdef LTC_OFB_MODE
int ofb_start(int cipher, const unsigned char *IV, const unsigned char *key,
              int keylen, int num_rounds, symmetric_OFB *ofb);
int ofb_encrypt(const unsigned char *pt, unsigned char *ct, unsigned long len, symmetric_OFB *ofb);
int ofb_decrypt(const unsigned char *ct, unsigned char *pt, unsigned long len, symmetric_OFB *ofb);
int ofb_getiv(unsigned char *IV, unsigned long *len, symmetric_OFB *ofb);
int ofb_setiv(const unsigned char *IV, unsigned long len, symmetric_OFB *ofb);
int ofb_done(symmetric_OFB *ofb);
#endif

#ifdef LTC_CBC_MODE
void cbc_start(int cipher, const unsigned char *IV, const unsigned char *key,
               int keylen, int num_rounds, symmetric_CBC *cbc);
void cbc_encrypt(const unsigned char *pt, unsigned char *ct, unsigned long len, symmetric_CBC *cbc);
void cbc_decrypt(const unsigned char *ct, unsigned char *pt, unsigned long len, symmetric_CBC *cbc);
void cbc_getiv(unsigned char *IV, unsigned long *len, symmetric_CBC *cbc);
void cbc_setiv(const unsigned char *IV, unsigned long len, symmetric_CBC *cbc);
void cbc_done(symmetric_CBC *cbc);
#endif

#ifdef LTC_CTR_MODE

#define CTR_COUNTER_LITTLE_ENDIAN    0x0000
#define CTR_COUNTER_BIG_ENDIAN       0x1000
#define LTC_CTR_RFC3686              0x2000

int ctr_start(               int   cipher,
              const unsigned char *IV,
              const unsigned char *key,       int keylen,
                             int  num_rounds, int ctr_mode,
                   symmetric_CTR *ctr);
int ctr_encrypt(const unsigned char *pt, unsigned char *ct, unsigned long len, symmetric_CTR *ctr);
int ctr_decrypt(const unsigned char *ct, unsigned char *pt, unsigned long len, symmetric_CTR *ctr);
int ctr_getiv(unsigned char *IV, unsigned long *len, symmetric_CTR *ctr);
int ctr_setiv(const unsigned char *IV, unsigned long len, symmetric_CTR *ctr);
int ctr_done(symmetric_CTR *ctr);
int ctr_test(void);
#endif

#ifdef LTC_LRW_MODE

#define LRW_ENCRYPT 0
#define LRW_DECRYPT 1

int lrw_start(               int   cipher,
              const unsigned char *IV,
              const unsigned char *key,       int keylen,
              const unsigned char *tweak,
                             int  num_rounds,
                   symmetric_LRW *lrw);
int lrw_encrypt(const unsigned char *pt, unsigned char *ct, unsigned long len, symmetric_LRW *lrw);
int lrw_decrypt(const unsigned char *ct, unsigned char *pt, unsigned long len, symmetric_LRW *lrw);
int lrw_getiv(unsigned char *IV, unsigned long *len, symmetric_LRW *lrw);
int lrw_setiv(const unsigned char *IV, unsigned long len, symmetric_LRW *lrw);
int lrw_done(symmetric_LRW *lrw);
int lrw_test(void);

/* don't call */
int lrw_process(const unsigned char *pt, unsigned char *ct, unsigned long len, int mode, symmetric_LRW *lrw);
#endif

#ifdef LTC_F8_MODE
int f8_start(                int  cipher, const unsigned char *IV,
             const unsigned char *key,                    int  keylen,
             const unsigned char *salt_key,               int  skeylen,
                             int  num_rounds,   symmetric_F8  *f8);
int f8_encrypt(const unsigned char *pt, unsigned char *ct, unsigned long len, symmetric_F8 *f8);
int f8_decrypt(const unsigned char *ct, unsigned char *pt, unsigned long len, symmetric_F8 *f8);
int f8_getiv(unsigned char *IV, unsigned long *len, symmetric_F8 *f8);
int f8_setiv(const unsigned char *IV, unsigned long len, symmetric_F8 *f8);
int f8_done(symmetric_F8 *f8);
int f8_test_mode(void);
#endif

#ifdef LTC_XTS_MODE
typedef struct {
   symmetric_key  key1, key2;
   int            cipher;
} symmetric_xts;

int xts_start(                int  cipher,
              const unsigned char *key1,
              const unsigned char *key2,
                    unsigned long  keylen,
                              int  num_rounds,
                    symmetric_xts *xts);

int xts_encrypt(
   const unsigned char *pt, unsigned long ptlen,
         unsigned char *ct,
         unsigned char *tweak,
         symmetric_xts *xts);
int xts_decrypt(
   const unsigned char *ct, unsigned long ptlen,
         unsigned char *pt,
         unsigned char *tweak,
         symmetric_xts *xts);

void xts_done(symmetric_xts *xts);
int  xts_test(void);
void xts_mult_x(unsigned char *I);
#endif

int find_cipher(const char *name);
int find_cipher_any(const char *name, int blocklen, int keylen);
int find_cipher_id(unsigned char ID);
int register_cipher(const struct ltc_cipher_descriptor *cipher);
int unregister_cipher(const struct ltc_cipher_descriptor *cipher);
bool cipher_is_valid(int idx);

LTC_MUTEX_PROTO(ltc_cipher_mutex)
