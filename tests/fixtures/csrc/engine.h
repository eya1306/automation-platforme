#ifndef ENGINE_H
#define ENGINE_H

#define MAX_SPEED   240
#define MIN_SPEED   0

typedef unsigned char  uint8;
typedef unsigned short uint16;

extern uint16 g_engine_rpm;

uint8 Engine_Update(uint16 rpm, uint8 gear);
void  Engine_Reset(void);

#endif
