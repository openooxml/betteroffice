#ifndef FONTPACK_STRING_H
#define FONTPACK_STRING_H
#include "types.h"
void *memcpy(void *dst, const void *src, size_t len);
void *memmove(void *dst, const void *src, size_t len);
void *memset(void *dst, int value, size_t len);
int memcmp(const void *left, const void *right, size_t len);
#endif
