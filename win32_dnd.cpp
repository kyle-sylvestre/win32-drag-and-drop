#include <Windows.h>

// from dropsource.cpp
extern HRESULT CreateDropSource(IDropSource **ppDropSource);

// from dataobject.cpp
extern HRESULT CreateDataObject(FORMATETC *fmtetc, STGMEDIUM *stgmeds, UINT count, IDataObject **ppDataObject);

// Drag one or more files to another window
// If you call CreateWindowEx with WS_EX_ACCEPTFILES, then do the following to prevent dropping onto yourself:
// DragAcceptFiles(your_window_handle, FALSE);
// Win32DragDrop(paths, num_paths);
// DragAcceptFiles(your_window_handle, TRUE);
void Win32DragDrop(const wchar_t **paths, int num_paths)
{
    IDataObject   *pDataObject = NULL;
    IDropSource   *pDropSource = NULL;
    HGLOBAL        hgDrop = NULL;
    DROPFILES*     pDrop = NULL;
    SIZE_T         uBuffSize = 0;
    SIZE_T         uTotalWideChars = 0;
    FORMATETC      etc = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };

    // Initialize OLE for using DoDragDrop function
    if (SUCCEEDED(OleInitialize(NULL)))
    {
        // Get the total number of wide characters needed
        for (int i = 0; i < num_paths; i++)
        {
            uTotalWideChars += lstrlenW(paths[i]) + 1;
        }

        // DROPFILES struct + null terminated wide strings + final L'\0'
        uBuffSize = sizeof(DROPFILES) + sizeof(wchar_t) * (uTotalWideChars + 1);

        // Allocate memory from the heap for the DROPFILES struct.
        if (hgDrop = GlobalAlloc(GHND | GMEM_SHARE, uBuffSize))
        {
            if (pDrop = (DROPFILES *)GlobalLock(hgDrop))
            {
                // Fill in the DROPFILES struct.
                pDrop->pFiles = sizeof(DROPFILES);

                // Indicate it contains Unicode strings.
                pDrop->fWide = TRUE;

                // Copy all the filenames into memory after the end of the DROPFILES struct.
                wchar_t *iter = (wchar_t *)(LPBYTE(pDrop) + sizeof(DROPFILES));
                for (int i = 0; i < num_paths; i++)
                {
                    int pathlen = lstrlenW(paths[i]);
                    CopyMemory(iter, paths[i], pathlen * sizeof(wchar_t));
                    iter += pathlen;
                    *iter++ = L'\0';
                }
                *iter++ = L'\0';
                GlobalUnlock(hgDrop);

                STGMEDIUM stg = {};
                stg.tymed = CF_HDROP;
                stg.hGlobal = hgDrop;

                CreateDropSource(&pDropSource);
                CreateDataObject(&etc, &stg, 1, &pDataObject);
                if (pDropSource && pDataObject)
                {
                    // do the drag/drop operation
                    // DROPEFFECT options from MSDN:
                    // DROPEFFECT_COPY: Drop results in a copy. The original data is untouched by the drag source.
                    // DROPEFFECT_MOVE: Drag source should remove the data.
                    // DROPEFFECT_LINK: Drag source should create a link to the original data.
                    DWORD dwEffect = 0;
                    DoDragDrop(pDataObject, pDropSource, DROPEFFECT_COPY, &dwEffect);
                }
                if (pDropSource) { pDropSource->Release(); pDropSource = NULL; }
                if (pDataObject) { pDataObject->Release(); pDataObject = NULL; }
            }
            GlobalFree(hgDrop); hgDrop = NULL;
        }

        OleUninitialize();
    }
}
