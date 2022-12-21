struct PostReceiveParam
{
    ATL::CInterfaceArray<IPhotoAcquireItem> sapAcquireItems;
    IPhotoAcquireSource* pPhotoAcquireSource;
    HRESULT hr;
    LPVOID pClass;
};

private:
    // Internal methods
    HRESULT TransferItem(IPhotoAcquireSource* pPhotoAcquireSource, IPhotoAcquireItem* pPhotoAcquireItem);
    static DWORD WINAPI  WorkerThread(LPVOID pVoid);

    CAtlList<PostReceiveParam*>                m_PostReceiveParamList;
    CONDITION_VARIABLE                         m_BufferNotEmpty;
    CONDITION_VARIABLE                         m_BufferNothingLeft;
    CONDITION_VARIABLE                         m_ErrorRecovered;
    CRITICAL_SECTION                           m_BufferLock;
    CRITICAL_SECTION                           m_ErrorLock;
    CRITICAL_SECTION                           m_PhotoAcqLock;
    BOOL                                                 m_bCloseWorkThread;
    PostReceiveParam*                          m_pErrorItem;

    HRESULT                                    m_hErrorFromWorkerThread;


DWORD WINAPI PhotoAcquire::WorkerThread(LPVOID pVoid)
{
    PhotoAcquire* RealThis = (PhotoAcquire*)pVoid;

    VerifyHrSucceeded(CoInitializeEx(0, COINIT_APARTMENTTHREADED));

    // We need to set up the thumbnail cache here (rather than in InitializeSession) because we can't access the thumbnail from a different
    // thread than the one that created the cache.
    RealThis->m_spThumbnailCache = NULL;
    //CComPtr<IThumbnailCache> spNewThumbnailCache = NULL;
    CComPtr<IThumbnailCache> spNewThumbnailCache = RealThis->m_spThumbnailCache;

    // Don't worry unduly about failures -- we just won't be able to pre-populate
    // the thumbnail cache (and we will check this interface before we use it)
    CoCreateInstance(CLSID_LocalThumbnailCache, NULL, CLSCTX_SERVER, IID_PPV_ARGS(&spNewThumbnailCache));

    // Again, don't worry about failures -- we just won't be able to index the imported photos
    RealThis->m_spSearchItemsChangedSink = NULL;
    CComPtr<ISearchManager> spSearchManager;
    if (SUCCEEDED(spSearchManager.CoCreateInstance(__uuidof(CSearchManager))))
    {
        // Get the catalog manager from the search manager
        CComPtr<ISearchCatalogManager> spCatalogManager;
        if (SUCCEEDED(spSearchManager->GetCatalog(L"SystemIndex", &spCatalogManager)))
        {
            // get the notification change sink so we can send notifications
            spCatalogManager->GetPersistentItemsChangedSink(&(RealThis->m_spSearchItemsChangedSink));
        }
    }
    BOOL bInCriticalSection=FALSE;

    // This loop does all the work of waiting for items and consuming them
    while(true)
    {
        try{
            EnterCriticalSection(&(RealThis->m_BufferLock));
            while((RealThis->m_PostReceiveParamList).IsEmpty())
            {
                // Terminate the thread after transfer or cancel
                 if (RealThis->m_bCloseWorkThread)
                 {
                    break;
                 }
                SleepConditionVariableCS(&(RealThis->m_BufferNotEmpty), &(RealThis->m_BufferLock), INFINITE);
                //WaitForMultipleObjects(2, , FALSE, INFINITE);
            }
            // Terminate the thread After transfer or cancel
            if ((RealThis->m_PostReceiveParamList).IsEmpty())
            {
                break;
            }
            // We added photos to the tail of the list, so we need to remove from the head to make sure everything stays in the right order.
            PostReceiveParam* pPRP = (RealThis->m_PostReceiveParamList).RemoveHead();
            LeaveCriticalSection(&(RealThis->m_BufferLock));


            EnterCriticalSection(&(RealThis->m_PhotoAcqLock));
            bInCriticalSection=TRUE;

            ATL::CInterfaceArray<IPhotoAcquireItem> *psapAcquireItems=&(pPRP->sapAcquireItems);
            IPhotoAcquireSource* pPhotoAcquireSource=pPRP->pPhotoAcquireSource;
            HRESULT hr=pPRP->hr;
            PhotoAcquire *pPA= (PhotoAcquire *)(pPRP->pClass);

            // Cache the target IPropertyStore for the temp item
            PROPVARIANT pvBoolTrue = {VT_BOOL, 0};
            VerifyHrSucceeded((*psapAcquireItems)[0]->SetProperty(PKEY_PhotoAcquire_SetTargetPropertyStoreCache, &pvBoolTrue));

            if (hr == S_OK)
            {
                for (size_t nCurrentItem = 0; nCurrentItem < (*psapAcquireItems).GetCount(); ++nCurrentItem)
                {
                    // Get this item's temp filename
                    Base::String strCurrentItemIntermediateFile;
                    VerifyHrOrThrow(pPA->GetItemFilename((*psapAcquireItems)[nCurrentItem], PKEY_PhotoAcquire_IntermediateFile, &strCurrentItemIntermediateFile));

                    // This is going to fail on XP, and isn't considered fatal in any case
                    CComPtr<IPropertyStore> spPropertiesToSave;
                    if (SUCCEEDED(pPA->CreateInMemoryPropertyStore(&spPropertiesToSave)))
                    {
                        // Get properties we intend to write to the file
                        if (VerifyHrSucceeded(pPA->GetWritableProperties(pPhotoAcquireSource, (*psapAcquireItems)[nCurrentItem], spPropertiesToSave)))
                        {
                            // Call the plugins before saving the file
                            if (pPA->m_PhotoAcquirePlugins.GetCount() != 0)
                            {
                                // Reopen the stream read-only
                                CComPtr<IStream> spStream;
                                if (VerifyHrSucceeded(SHCreateStreamOnFileEx(strCurrentItemIntermediateFile, STGM_READ|STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, FALSE, NULL, &spStream)))
                                {
                                    // Call the plugins
                                    VerifyHrSucceeded(pPA->CallPlugins(PAPS_PRESAVE, (*psapAcquireItems)[nCurrentItem], spStream, NULL, spPropertiesToSave));
                                }
                            }

                            // Write the properties to the temp file                    
                            if (!Helpers::IsTransferFlagSet(pPhotoAcquireSource, PHOTOACQ_DISABLE_METADATA_WRITE))
                            {
                                pPA->WritePostAcquirePropertiesToFinalFile(spPropertiesToSave, strCurrentItemIntermediateFile);
                            }
                        }
                    }
                }

                for (size_t nCurrentItem = 0; nCurrentItem < (*psapAcquireItems).GetCount(); ++nCurrentItem)
                {
                    // Get this item's temp filename
                    Base::String strCurrentItemIntermediateFile;
                    VerifyHrOrThrow(pPA->GetItemFilename((*psapAcquireItems)[nCurrentItem], PKEY_PhotoAcquire_IntermediateFile, &strCurrentItemIntermediateFile));

                    // Save it and get the final filename
                    WCHAR szFinalFilename[MAX_PATH] = L"";
                    VerifyHrOrThrow(pPA->SaveItemToDisk(pPhotoAcquireSource, (*psapAcquireItems)[nCurrentItem], nCurrentItem == 0, szFinalFilename, ARRAYSIZE(szFinalFilename)));

                    // Call the plugins
                    VerifyHrSucceeded(pPA->CallPlugins(PAPS_POSTSAVE, (*psapAcquireItems)[nCurrentItem], NULL, szFinalFilename, NULL));

                    // Set the filetimes.  Don't care if this fails, all that happens is the date created/modified of the file is off.
                    PHOTOACQ_IGNORE_RESULT(pPA->UpdateFiletimes((*psapAcquireItems)[nCurrentItem], szFinalFilename));

                    // Make sure the thumbnail is populated if this feature is enabled.  Don't care if this fails, all that happens is
                    // the thumbnail is not prepopulated.
                    if (Helpers::IsTransferFlagSet(pPhotoAcquireSource, PHOTOACQ_ENABLE_THUMBNAIL_CACHING))
                    {
                        PHOTOACQ_IGNORE_RESULT(pPA->PrePopulateThumbnail(szFinalFilename));
                    }

                    // Reset the file attributes so this file is visible to the gallery database and MS Search
                    // This is a fatal error, because we don't want to leave hidden files laying around.
                    VerifyHrOrThrow(pPA->SetFileAttributesWithRetry(szFinalFilename, 0, FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_NOT_CONTENT_INDEXED, kdwSetFileAttributesRetryWait, kdwSetFileAttributesRetryCount));

                    // Add it to duplicate tracking
                    ATL::CComQIPtr<IPhotoAcquireDuplicateTrackingSession> spDuplicateSession(pPhotoAcquireSource);
                    if (spDuplicateSession != NULL)
                    {
                        VerifyHrSucceeded(spDuplicateSession->AddItemToDuplicatesList((*psapAcquireItems)[nCurrentItem]));
                    }
                }
            }

            // Remove all temp files
            PROPVARIANT propVariantNull = {VT_NULL, 0};
            for (size_t nCurrentItem = 0; nCurrentItem < (*psapAcquireItems).GetCount(); ++nCurrentItem)
            {
                // Get this item's temp filename
                Base::String strCurrentItemIntermediateFile;
                if (SUCCEEDED(pPA->GetItemFilename((*psapAcquireItems)[nCurrentItem], PKEY_PhotoAcquire_IntermediateFile, &strCurrentItemIntermediateFile)))
                {
                    // Delete the temp file
                    pPA->DeleteFileWithRetry(strCurrentItemIntermediateFile, kdwDeleteRetryWait, knDeleteRetryCount);

                    // Zero out the temp file entry
                    VerifyHrSucceeded((*psapAcquireItems)[nCurrentItem]->SetProperty(PKEY_PhotoAcquire_IntermediateFile, &propVariantNull));
                }

                if (nCurrentItem == 0)
                {
                    // Uncache the IPropertyStore for the source and target items
                    PROPVARIANT pvBoolFalse = {VT_BOOL, 0};
                    pvBoolFalse.boolVal = VARIANT_FALSE;
                    VerifyHrSucceeded((*psapAcquireItems)[0]->SetProperty(PKEY_PhotoAcquire_SetSourcePropertyStoreCache, &pvBoolFalse));
                    VerifyHrSucceeded((*psapAcquireItems)[0]->SetProperty(PKEY_PhotoAcquire_SetTargetPropertyStoreCache, &pvBoolFalse));
                }
            }

            // If we skipped this file because it was transcoded for sync, we want to set an
            // error on the items, but continue to return success
            PROPVARIANT propVariantResult = {0};
            propVariantResult.vt = VT_ERROR;
            propVariantResult.scode = (hr == S_FALSE) ? E_FAIL : hr;
            for (size_t nCurrentItem = 0; nCurrentItem < (*psapAcquireItems).GetCount(); ++nCurrentItem)
            {
                // Set the result
                VerifyHrSucceeded((*psapAcquireItems)[nCurrentItem]->SetProperty(PKEY_PhotoAcquire_TransferResult, &propVariantResult));
            }
            delete (pPRP);
            LeaveCriticalSection(&(RealThis->m_PhotoAcqLock));
            bInCriticalSection=FALSE;

            EnterCriticalSection(&(RealThis->m_BufferLock));
            if((RealThis->m_PostReceiveParamList).IsEmpty())
            {
                LeaveCriticalSection(&(RealThis->m_BufferLock));
                WakeConditionVariable(&(RealThis->m_BufferNothingLeft));
            }
            else
            {
                LeaveCriticalSection(&(RealThis->m_BufferLock));
            }
        }
        catch(Base::Exception e)
        {
            if (TRUE == bInCriticalSection)
            {
                LeaveCriticalSection(&(RealThis->m_PhotoAcqLock));
            }
            
            // Set an error somewhere in RealThis that the producer thread can access later on.
            EnterCriticalSection(&(RealThis->m_ErrorLock));
            RealThis->m_hErrorFromWorkerThread = e;
            SleepConditionVariableCS(&(RealThis->m_ErrorRecovered), &(RealThis->m_ErrorLock), INFINITE);
            RealThis->m_hErrorFromWorkerThread = S_OK;
            LeaveCriticalSection(&(RealThis->m_ErrorLock));
        }
    }

    return 0;
}

HRESULT PhotoAcquire::TransferItem(IPhotoAcquireSource* pPhotoAcquireSource, IPhotoAcquireItem* pPhotoAcquireItem)
{
    AutoPerfTrace(TRACEID_PHOTOACQUIRE_TRANSFER_ITEM);
    AUTO_FUNCTION_TIMER();
    HRESULT hr = S_OK;
    PostReceiveParam *pPRP = new PostReceiveParam();
    try
    {
        // Construct an array of items consisting of the main item and any attachments
        ATL::CInterfaceArray<IPhotoAcquireItem> sapAcquireItems;
        VerifyHrOrThrow(CreateItemArray(pPhotoAcquireItem, &sapAcquireItems));

        // This inner try-catch block is here so we can handle any errors by cleaning
        // up the temp files and final files for this group of related items.
        try
        {
            // Loop through all of the items and copy them to the temp directory
            for (size_t nCurrentItem = 0; nCurrentItem < sapAcquireItems.GetCount(); ++nCurrentItem)
            {
                // If this is the first item, and it is marked "skip", we won't continue
                if (nCurrentItem == 0 && Helpers::IsMarkedSkip(sapAcquireItems[0]))
                {
                    // Set the cancelled HRESULT and break out of the loop.
                    hr = S_FALSE;
                    break;
                }

                // Cache the source IPropertyStore for the primary item
                PROPVARIANT pvBoolTrue = {VT_BOOL, 0};
                pvBoolTrue.boolVal = VARIANT_TRUE;
                if (nCurrentItem == 0)
                {
                    VerifyHrSucceeded(sapAcquireItems[0]->SetProperty(PKEY_PhotoAcquire_SetSourcePropertyStoreCache, &pvBoolTrue));
                }

                // Read item and save it to a temp file.
                VerifyHrOrThrowSilentFor(E_ABORT, ReadItemFromDevice(sapAcquireItems[nCurrentItem], nCurrentItem == 0 ? m_spPhotoAcquireProgressCB : NULL));

                // Don't continue if the main item was transcoded for sync
                if (nCurrentItem == 0 && Helpers::IsTranscodedForSync(sapAcquireItems[0]))
                {
                    //
                    // 1) Add it to duplicates DB so that we don't enumerate it in the future
                    // 2) Increment our "skipped transcoded items" count (drives displaying the dialog at the end of acquisition)
                    // 3) Set the cancelled HRESULT and break out of the loop.
                    //

                    ATL::CComQIPtr<IPhotoAcquireDuplicateTrackingSession> spDuplicateSession(pPhotoAcquireSource);
                    if (spDuplicateSession != NULL)
                    {
                        spDuplicateSession->AddItemToDuplicatesList(sapAcquireItems[nCurrentItem]);
                    }
                    m_dwSkippedTranscodedItems++;

                    hr = S_FALSE;
                    break;
                }
            }
        }
        catch (Base::Exception e)
        {
            if(pPRP == NULL)
            {
                delete(pPRP);
            }
            hr = e;
        }

        // If an error occurred, remove all final files
        if (FAILED(hr))
        {
            for (size_t nCurrentItem = 0; nCurrentItem < sapAcquireItems.GetCount(); ++nCurrentItem)
            {
                // Get this item's final filename
                Base::String strCurrentItemFinalFile;
                if (SUCCEEDED(GetItemFilename(sapAcquireItems[nCurrentItem], PKEY_PhotoAcquire_FinalFilename, &strCurrentItemFinalFile)))
                {
                    // Tell the plugins we are cleaning up
                    VerifyHrSucceeded(CallPlugins(PAPS_CLEANUP, sapAcquireItems[nCurrentItem], NULL, strCurrentItemFinalFile, NULL));

                    // Delete the target file
                    DeleteFileWithRetry(strCurrentItemFinalFile, kdwDeleteRetryWait, knDeleteRetryCount);

                    // Remove this item from duplicate tracking.
                    ATL::CComQIPtr<IPhotoAcquireDuplicateTrackingSession> spDuplicateSession(pPhotoAcquireSource);
                    if (spDuplicateSession != NULL)
                    {
                         VerifyHrSucceeded(spDuplicateSession->RemoveItemFromDuplicatesList(sapAcquireItems[nCurrentItem]));
                    }
                    PROPVARIANT propVariantNull = {VT_NULL, 0};

                    // Zero out the target file entry
                    VerifyHrSucceeded(sapAcquireItems[nCurrentItem]->SetProperty(PKEY_PhotoAcquire_FinalFilename, &propVariantNull));
                }
            }
        }

        // Add item to the list of items wor the worker thread to process

        for (size_t nCurrentItem = 0; nCurrentItem < sapAcquireItems.GetCount(); ++nCurrentItem)
        {
            // Set the result
            pPRP->sapAcquireItems.Add(sapAcquireItems[nCurrentItem]);
        }

        pPRP->pPhotoAcquireSource = pPhotoAcquireSource;
        pPRP->hr=hr;
        pPRP->pClass= this;


        EnterCriticalSection(&m_BufferLock);
        // We add photos to the tail of our list, so to make sure that photos are processed in the same order they were imported, we should
        // make sure to process them from the head of the list.
        m_PostReceiveParamList.AddTail(pPRP);
        LeaveCriticalSection(&m_BufferLock);
        WakeConditionVariable(&m_BufferNotEmpty);
    }
    catch (Base::Exception e)
    {
        if (pPRP==NULL)
        {
            delete (pPRP);
        }
        hr = e;
    }

    return hr;
}

HRESULT PhotoAcquire::Transfer(IPhotoAcquireSource* pPhotoAcquireSource)
{
    AutoPerfTrace(TRACEID_PHOTOACQUIRE_TRANSFER_ALL_ITEMS);
    AUTO_FUNCTION_TIMER();
    HANDLE hThread = NULL;

    // Make sure we haven't been cancelled
    HRESULT hr = Helpers::IsCancelled(m_spPhotoAcquireProgressCB);
    if (SUCCEEDED(hr))
    {
        InitializeConditionVariable(&m_BufferNotEmpty);
        InitializeConditionVariable(&m_BufferNothingLeft);
        InitializeConditionVariable(&m_ErrorRecovered);
        InitializeCriticalSection(&m_BufferLock);
        InitializeCriticalSection(&m_ErrorLock);
        InitializeCriticalSection(&m_PhotoAcqLock);
        EnterCriticalSection(&m_PhotoAcqLock);
        m_bCloseWorkThread=FALSE;
        // Create our worker thread
        DWORD idThread;
        hThread = CreateThread (NULL,
            0,
            WorkerThread,
            this,
            0,
            &(idThread));
        if(hThread == NULL)
        {
            hr = E_FAIL;
        }
        if (VerifyHrSucceeded(hr))
        {
            // Tell the callback we are starting the transfer
            hr = m_spPhotoAcquireProgressCB->StartTransfer(pPhotoAcquireSource);
            if (VerifyHrSucceeded(hr))
            {
                UINT nAcquireItemCount = 0;
                hr = pPhotoAcquireSource->GetItemCount(&nAcquireItemCount);
                if (VerifyHrSucceeded(hr))
                {
                    // Loop through our list of items
                    UINT nCurrentItem = 0;
                    while (SUCCEEDED(hr) && nCurrentItem < nAcquireItemCount)
                    {
                        // Make sure that our worker thread hasn't run into any errors so far
                        // If the worker thread has run into a problem, we'll display an error message.  We need to handle this separately from
                        // errors coming out of TransferItem() because the only access we have to what the worker thread is doing is through
                        // the queue of items to import.
                        EnterCriticalSection(&m_ErrorLock);
                        hr = m_hErrorFromWorkerThread;
                        LeaveCriticalSection(&m_ErrorLock);
                        if(FAILED(hr))
                        {
                            ERROR_ADVISE_RESULT nErrorAdviseResult;
                            if (VerifyHrSucceeded(DisplayErrorMessage(hr, (m_pErrorItem->sapAcquireItems)[0], IDS_DEFAULT_TRANSFER_ERROR, IDS_DEFAULT_TRANSFER_ERROR_NAMED, PHOTOACQUIRE_ERROR_SKIPRETRYCANCEL, &nErrorAdviseResult)))
                            {
                                switch (nErrorAdviseResult)
                                {
                                case PHOTOACQUIRE_RESULT_SKIP:
                                case PHOTOACQUIRE_RESULT_SKIP_ALL:
                                    // Skip this item
                                    hr = S_OK;
                                    EnterCriticalSection(&m_ErrorLock);
                                    m_hErrorFromWorkerThread = S_OK;
                                    LeaveCriticalSection(&m_ErrorLock);
                                    WakeConditionVariable(&m_ErrorRecovered);
                                    break;

                                case PHOTOACQUIRE_RESULT_RETRY:
                                    // Retry this item by re-adding it to the head of the list (so it will be the first thing the worker thread processes)
                                    EnterCriticalSection(&m_BufferLock);
                                    m_PostReceiveParamList.AddHead(m_pErrorItem);
                                    LeaveCriticalSection(&m_BufferLock);
                                    hr = S_OK;
                                    EnterCriticalSection(&m_ErrorLock);
                                    m_hErrorFromWorkerThread = S_OK;
                                    LeaveCriticalSection(&m_ErrorLock);
                                    WakeConditionVariable(&m_ErrorRecovered);
                                    break;

                                case PHOTOACQUIRE_RESULT_ABORT:
                                    // Exit the loop, and set the cancelled state
                                    // Do we want to clear the list of items to process as well?  Or do we want to let those items process first?
                                    hr = E_ABORT;
                                    break;
                                }
                            }
                        }
                        if(VerifyHrSucceeded(hr))
                        {

                            // Get the next item
                            CComPtr<IPhotoAcquireItem> spPhotoAcquireItem;
                            hr = pPhotoAcquireSource->GetItemAt(nCurrentItem, &spPhotoAcquireItem);
                            if (VerifyHrSucceeded(hr))
                            {
                                // Have we been cancelled?
                                hr = Helpers::IsCancelled(m_spPhotoAcquireProgressCB);
                                if (SUCCEEDED(hr))
                                {
                                    // Tell the callback what percentage done we are
                                    hr = m_spPhotoAcquireProgressCB->UpdateTransferPercent(TRUE, MulDiv(nCurrentItem, 100, nAcquireItemCount));
                                    if (VerifyHrSucceeded(hr))
                                    {
                                        // Tell the callback we are starting the transfer of this item
                                        hr = m_spPhotoAcquireProgressCB->StartItemTransfer(nCurrentItem, spPhotoAcquireItem);
                                        if (VerifyHrSucceeded(hr))
                                        {
                                            // Transfer it
                                            hr = TransferItem(pPhotoAcquireSource, spPhotoAcquireItem);

                                            // We have to tell the client we are done, even if it failed.
                                            // BUT! We don't want to overwrite an error with S_OK, so we
                                            // only save the result of EndItemTransfer if no errors had
                                            // occurred AND we get an error from the callback.
                                            HRESULT hrEnd = m_spPhotoAcquireProgressCB->EndItemTransfer(nCurrentItem, spPhotoAcquireItem, hr);
                                            if (SUCCEEDED(hr) && FAILED(hrEnd))
                                            {
                                                hr = hrEnd;
                                            }
                                        }
                                    }
                                }
                            }

                            if (FAILED(hr) && hr != E_ABORT)
                            {
                                // Display an error message.  If an error occurs displaying the error message, just ignore it.
                                // In this case, we will be returning the original error.
                                ERROR_ADVISE_RESULT nErrorAdviseResult;
                                if (VerifyHrSucceeded(DisplayErrorMessage(hr, spPhotoAcquireItem, IDS_DEFAULT_TRANSFER_ERROR, IDS_DEFAULT_TRANSFER_ERROR_NAMED, PHOTOACQUIRE_ERROR_SKIPRETRYCANCEL, &nErrorAdviseResult)))
                                {
                                    switch (nErrorAdviseResult)
                                    {
                                    case PHOTOACQUIRE_RESULT_SKIP:
                                    case PHOTOACQUIRE_RESULT_SKIP_ALL:
                                        // Skip this item
                                        // TODO: Implement "Skip All" errors
                                        hr = S_OK;
                                        ++nCurrentItem;
                                        break;

                                    case PHOTOACQUIRE_RESULT_RETRY:
                                        // Retry this item
                                        hr = S_OK;
                                        break;

                                    case PHOTOACQUIRE_RESULT_ABORT:
                                        // Exit the loop and set the cancelled state
                                        // TODO: we also need to kill the worker thread.  Do we want to allow it to finish the remaining
                                        // items it has?
                                        // Finish the remaining itens and leave
                                        LeaveCriticalSection(&m_PhotoAcqLock);
                                        m_bCloseWorkThread=TRUE;
                                        WakeConditionVariable(&m_BufferNotEmpty);
                                        EnterCriticalSection(&m_BufferLock);
                                        SleepConditionVariableCS(&m_BufferNothingLeft, &m_BufferLock, INFINITE); // don't use events here
                                        LeaveCriticalSection(&m_BufferLock);
                                        hr = E_ABORT;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                ++nCurrentItem;

                                // When we get to the very end, we need to make sure to give the worker thread a chance to 
                                // finish working
                                if (nCurrentItem == nAcquireItemCount)
                                {
                                    LeaveCriticalSection(&m_PhotoAcqLock);     
                                    m_bCloseWorkThread=TRUE;
                                    WakeConditionVariable(&m_BufferNotEmpty);
                                    EnterCriticalSection(&m_BufferLock);
                                    SleepConditionVariableCS(&m_BufferNothingLeft, &m_BufferLock, INFINITE); // don't use events here
                                    LeaveCriticalSection(&m_BufferLock);
                                }
                            }
                        }
                    }
                }
                else if (hr != E_ABORT)
                {
                    DisplayErrorMessage(hr, NULL, IDS_UNABLE_TO_START_TRANSFER, 0, PHOTOACQUIRE_ERROR_OK, NULL, false);
                }

                // Tell the callback we are done acquiring
                PHOTOACQ_IGNORE_RESULT(m_spPhotoAcquireProgressCB->EndTransfer(hr));
            }
            else if (hr != E_ABORT)
            {
                DisplayErrorMessage(hr, NULL, IDS_UNABLE_TO_START_TRANSFER, 0, PHOTOACQUIRE_ERROR_OK, NULL, false);
            }
        }
    }
    // signla worker thread to exit
    // wait for single object (hthread, 2-3 seconds)
    
    if(hThread != NULL)
    {
         m_bCloseWorkThread=TRUE;
    //WaitForSingleObject(hthread, 2000);
    }
    return hr;
}


