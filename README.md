# GetVideoMemory
A small utility to display total video RAM

Unlike system memory, getting the amount of video memory isn't straightforward. The biggest issue people run into is the well known methods use a size_t variable for it, which is only 4 bytes for 32bit applications therefore can't represent more than 4GB, which many video cards have these days. Windows 10/11 have other options, but it's possible back to 7/Vista.

Workarounds typically given are awkward and hard to find, sometimes referring to very old DX versions. This method is more up to date, compatible, and reliable. It will display the amount of memory each adapter on the system has (given by ID and description like e.g. "NVIDIA GeForce Rtx 4070 Ti", and note which one displays the primary desktop.


```vba
'Roughly based on code from Pavel Yosifovich's DeviceExplorer
Private Sub DisplayVideoMem()
    On Error Resume Next
    Dim pFactory As IDXGIFactory
    Dim pAdapters As D3DKMT_ENUMADAPTERS
    Dim adapter As IDXGIAdapter
    Dim desc As DXGI_ADAPTER_DESC
    Dim status As NTSTATUS
    Dim sDesc As String
    Dim segInfo As D3DKMT_SEGMENTSIZEINFO
    Dim n As Long
    Dim hr As Long = CreateDXGIFactory(IID_IDXGIFactory, pFactory)
    If SUCCEEDED(hr) Then
        status = D3DKMTEnumAdapters(pAdapters)
        For i As Long = 0 To pAdapters.NumAdapters - 1
            n = -1
            pFactory.EnumAdapters(i, adapter)
            If SUCCEEDED(Err.LastHresult) Then
                adapter.GetDesc(desc)
            End If
            For j As Long = 0 To pAdapters.NumAdapters - 1
                If CompareMemory(desc.AdapterLuid, pAdapters.Adapters(j).AdapterLuid, LenB(Of LUID)) = LenB(Of LUID) Then
                    n = j
                    Exit For
                End If
            Next
            If n >= 0 Then
                sDesc = WCHARtoSTR(desc.Description)
                If QueryAdapterInfo(pAdapters.Adapters(n).hAdapter, KMTQAITYPE_GETSEGMENTSIZE, segInfo) = STATUS_SUCCESS Then
                    AppendLog If(i = 0, "*", "") & "Adapter " & CStr(i + 1) & " (id=" & desc.deviceId & "): " & sDesc & ", " & segInfo.DedicatedVideoMemorySize & " bytes"
                End If
            End If
            Set adapter = Nothing
            ZeroMemory desc, LenB(desc)
            ZeroMemory segInfo, LenB(segInfo)
            sDesc = ""
        Next
    End If
End Sub

Private Function QueryAdapterInfo(Of T)(ByVal hAdapter As Long, ByVal type As KMTQUERYADAPTERINFOTYPE, data As T) As Long
    Dim queryInfo As D3DKMT_QUERYADAPTERINFO
    queryInfo.hAdapter = hAdapter
    queryInfo.Type = type
    queryInfo.pPrivateDriverData = VarPtr(data)
    queryInfo.PrivateDriverDataSize = LenB(data)
    Dim status As Long = D3DKMTQueryAdapterInfo(queryInfo)
    Return status
End Function
```

Requires Windows Development Library for twinBASIC v9.1.578+.
It's linked not embedded... if you don't have it in the linked packages folder, or have an older version, uncheck the existing reference then re-add from the Available packages tab, then uncheck Embedded to export to the shared packages folder. 
