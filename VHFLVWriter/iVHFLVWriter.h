
#ifndef __iVHFLVWriter_h__
#define __iVHFLVWriter_h__

#ifdef __cplusplus
extern "C" {
#endif


// audio stream flags ( _VHFLV_stream_set::dwflags )

#define VHFLVWriter_ARate_5_5			0x00000000
#define VHFLVWriter_ARate_11			0x00000004
#define VHFLVWriter_ARate_22			0x00000008
#define VHFLVWriter_ARate_44			0x0000000C
#define VHFLVWriter_ARate_mask			0x0000000C

#define VHFLVWriter_AType_mono			0x00000000
#define VHFLVWriter_AType_ster			0x00000001
#define VHFLVWriter_AType_mask			0x00000001



// {AD53C322-17C4-41ce-99C4-FBBA1C685BC1}
DEFINE_GUID(IID_IVHFLVWriter,
0xad53c322, 0x17c4, 0x41ce, 0x99, 0xc4, 0xfb, 0xba, 0x1c, 0x68, 0x5b, 0xc1);


struct _VHFLV_stream_set
{
	DWORD bitrate; // kbps
	DWORD dwflags; // flags
};


DECLARE_INTERFACE_(IVHFLVWriter, IUnknown)
{

	STDMETHOD(get_vstream) (THIS_
		_VHFLV_stream_set *pstream
	) PURE;

	STDMETHOD(put_vstream) (THIS_
		_VHFLV_stream_set *pstream
	) PURE;

	STDMETHOD(get_astream) (THIS_
		_VHFLV_stream_set *pstream
	) PURE;

	STDMETHOD(put_astream) (THIS_
		_VHFLV_stream_set *pstream
	) PURE;

};

#ifdef __cplusplus
}
#endif

#endif // __iVHFLVWriter_h__
