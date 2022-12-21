HRESULT PhotoAcquire::TransferItem(IPhotoAcquireSource* pPhotoAcquireSource, IPhotoAcquireItem* pPhotoAcquireItem)
{
    AutoPerfTrace(TRACEID_PHOTOACQUIRE_TRANSFER_ITEM);
    AUTO_FUNCTION_TIMER();
    HRESULT hr = S_OK;
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
                // We already checked for the skip case above; don't need to repeat it here.
                
                // Cache the source IPropertyStore for the primary item
                PROPVARIANT pvBoolTrue = {VT_BOOL, 0};
                pvBoolTrue.boolVal = VARIANT_TRUE;
                if (nCurrentItem == 0)
                {
                    VerifyHrSucceeded(sapAcquireItems[0]->SetProperty(PKEY_PhotoAcquire_SetSourcePropertyStoreCache, &pvBoolTrue));
                }

                // Read item and save it to a temp file.
                VerifyHrOrThrowSilentFor(E_ABORT, ReadItemFromDevice(sapAcquireItems[nCurrentItem], nCurrentItem == 0 ? m_spPhotoAcquireProgressCB : NULL));

                // Cache the target IPropertyStore for the temp item
                if (nCurrentItem == 0)
                {
                    VerifyHrSucceeded(sapAcquireItems[0]->SetProperty(PKEY_PhotoAcquire_SetTargetPropertyStoreCache, &pvBoolTrue));
                }

                // Don't continue if the main item was transcoded for sync
                if (nCurrentItem == 0 && AcquisitionHelpers::IsTranscodedForSync(sapAcquireItems[0]))
                {
                    // Set the cancelled HRESULT and break out of the loop.
                    hr = S_FALSE;
                    break;
                }
            }

            // Make sure we didn't cancel
            if (hr == S_OK)
            {
                for (size_t nCurrentItem = 0; nCurrentItem < sapAcquireItems.GetCount(); ++nCurrentItem)
                {
                    // Get this item's temp filename
                    Base::String strCurrentItemIntermediateFile;
                    VerifyHrOrThrow(GetItemFilename(sapAcquireItems[nCurrentItem], PKEY_PhotoAcquire_IntermediateFile, &strCurrentItemIntermediateFile));

                    CComPtr<IPropertyStore> spPropertiesToSave;
                    if (SUCCEEDED(CreateInMemoryPropertyStore(&spPropertiesToSave)))
                    {
                        // Get the tags we intend to write to the file
                        if (VerifyHrSucceeded(MergeAndAddTags(sapAcquireItems[nCurrentItem], spPropertiesToSave)))
                        {
                            // Call the plugins before saving the file
                            if (m_photoAcquirePlugins.GetCount() != 0)
                            {
                                // Reopen the stream read-only
                                CComPtr<IStream> spStream;
                                if (VerifyHrSucceeded(SHCreateStreamOnFileEx(strCurrentItemIntermediateFile, STGM_READ|STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, FALSE, NULL, &spStream)))
                                {
                                    // Call the plugins
                                    VerifyHrSucceeded(CallPlugins(PAPS_PRESAVE, sapAcquireItems[nCurrentItem], spStream, NULL, spPropertiesToSave));
                                }
                            }

                            // Write the properties to the temp file                    
                            if (!AcquisitionHelpers::IsTransferFlagSet(pPhotoAcquireSource, PHOTOACQ_DISABLE_METADATA_WRITE))
                            {
                                WritePostAcquirePropertiesToFinalFile(spPropertiesToSave, strCurrentItemIntermediateFile);
                            }
                        }
                    }
                }

                for (size_t nCurrentItem = 0; nCurrentItem < sapAcquireItems.GetCount(); ++nCurrentItem)
                {
                    // Get this item's temp filename
                    Base::String strCurrentItemIntermediateFile;
                    VerifyHrOrThrow(GetItemFilename(sapAcquireItems[nCurrentItem], PKEY_PhotoAcquire_IntermediateFile, &strCurrentItemIntermediateFile));

                    // Save it and get the final filename
                    WCHAR szFinalFilename[MAX_PATH] = L"";
                    VerifyHrOrThrow(SaveItemToDisk(pPhotoAcquireSource, sapAcquireItems[nCurrentItem], nCurrentItem == 0, szFinalFilename, ARRAYSIZE(szFinalFilename)));

                    // Call the plugins
                    VerifyHrSucceeded(CallPlugins(PAPS_POSTSAVE, sapAcquireItems[nCurrentItem], NULL, szFinalFilename, NULL));

                    // Set the filetimes
                    UpdateFiletimes(sapAcquireItems[nCurrentItem], szFinalFilename);

                    // Reset the file attributes so this file is visible to the gallery database and MS Search
                    // This is a fatal error, because we don't want to leave hidden files laying around.
                    VerifyHrOrThrow(SetFileAttributesWithRetry(szFinalFilename, 0, FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_NOT_CONTENT_INDEXED, kdwSetFileAttributesRetryWait, kdwSetFileAttributesRetryCount));

                    // Add it to the database
                    PhotoLibraryIntegration::Database::SupportedStatus eLibrarySupportedStatus = PhotoLibraryIntegration::Database::keLibraryUnsupported;
                    VerifyHrSucceeded(m_photoAcquireDatabaseIntegration.AddFileToDatabase(szFinalFilename, sapAcquireItems[nCurrentItem], &eLibrarySupportedStatus));
                    if (eLibrarySupportedStatus != PhotoLibraryIntegration::Database::keLibraryUnsupported)
                    {
                        m_fLibrarySupportedFiles = true;
                    }

                    // Add it to duplicate tracking
                    ATL::CComQIPtr<IPhotoAcquireDuplicateTrackingSession> spDuplicateSession(pPhotoAcquireSource);
                    if (spDuplicateSession != NULL)
                    {
                        VerifyHrSucceeded(spDuplicateSession->AddItemToDuplicatesList(sapAcquireItems[nCurrentItem]));
                    }
                }
            }
        }
        catch (Base::Exception e)
        {
           hr = e;
        }


        // Remove all temp files
        PROPVARIANT propVariantNull = {VT_NULL, 0};
        for (size_t nCurrentItem = 0; nCurrentItem < sapAcquireItems.GetCount(); ++nCurrentItem)
        {
            // Get this item's temp filename
            Base::String strCurrentItemIntermediateFile;
            if (SUCCEEDED(GetItemFilename(sapAcquireItems[nCurrentItem], PKEY_PhotoAcquire_IntermediateFile, &strCurrentItemIntermediateFile)))
            {
                // Delete the temp file
                DeleteFileWithRetry(strCurrentItemIntermediateFile, kdwDeleteRetryWait, knDeleteRetryCount);

                // Zero out the temp file entry
                VerifyHrSucceeded(sapAcquireItems[nCurrentItem]->SetProperty(PKEY_PhotoAcquire_IntermediateFile, &propVariantNull));
            }

            if (nCurrentItem == 0)
            {
                // Uncache the IPropertyStore for the source and target items
                PROPVARIANT pvBoolFalse = {VT_BOOL, 0};
                pvBoolFalse.boolVal = VARIANT_FALSE;
                VerifyHrSucceeded(sapAcquireItems[0]->SetProperty(PKEY_PhotoAcquire_SetSourcePropertyStoreCache, &pvBoolFalse));
                VerifyHrSucceeded(sapAcquireItems[0]->SetProperty(PKEY_PhotoAcquire_SetTargetPropertyStoreCache, &pvBoolFalse));
            }
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

                    // Delete it from the database
                    VerifyHrSucceeded(m_photoAcquireDatabaseIntegration.RemoveFileFromDatabase(sapAcquireItems[nCurrentItem]));

                    // Zero out the target file entry
                    VerifyHrSucceeded(sapAcquireItems[nCurrentItem]->SetProperty(PKEY_PhotoAcquire_FinalFilename, &propVariantNull));
                }
            }
        }

        // If we skipped this file because it was transcoded for sync, we want to set an
        // error on the items, but continue to return success
        PROPVARIANT propVariantResult = {0};
        propVariantResult.vt = VT_ERROR;
        propVariantResult.scode = (hr == S_FALSE) ? E_FAIL : hr;
        for (size_t nCurrentItem = 0; nCurrentItem < sapAcquireItems.GetCount(); ++nCurrentItem)
        {
            // Set the result
            VerifyHrSucceeded(sapAcquireItems[nCurrentItem]->SetProperty(PKEY_PhotoAcquire_TransferResult, &propVariantResult));
        }
    }
    catch (Base::Exception e)
    {
       hr = e;
    }

    return hr;
}

HRESULT PhotoAcquire::InitializeSession(HWND hWndParent, IPhotoAcquireProgressCB* pPhotoAcquireProgressCB, IPhotoAcquireSource* pPhotoAcquireSource)
{
    AUTO_FUNCTION_TIMER();
    // Make sure everything has been released
    m_spPhotoAcquireProgressCB = NULL;
    m_spPathnameFromTemplate = NULL;

    // Get a pointer to the cancel callback
    CComPtr<IPhotoProgressDialogCancelCB> spPhotoProgressDialogCancelCB;
    HRESULT hr = this->QueryInterface(IID_PPV_ARGS(&spPhotoProgressDialogCancelCB));
    if (VerifyHrSucceeded(hr))
    {
        // Create our callback
        hr = DefaultPhotoAcquireProgressCB_CreateInstance(pPhotoAcquireProgressCB, pPhotoAcquireSource, hWndParent, spPhotoProgressDialogCancelCB, &m_spPhotoAcquireProgressCB);
        if (VerifyHrSucceeded(hr))
        {
            // Create the pathname generator
            hr = PathnameFromTemplate_CreateInstance(&m_spPathnameFromTemplate);
        }
    }

    return hr;
}

HRESULT PhotoAcquire::Transfer(IPhotoAcquireSource* pPhotoAcquireSource)
{
    AutoPerfTrace(TRACEID_PHOTOACQUIRE_TRANSFER_ALL_ITEMS);
    AUTO_FUNCTION_TIMER();

    // Make sure we haven't been cancelled
    HRESULT hr = AcquisitionHelpers::IsCancelled(m_spPhotoAcquireProgressCB);
    if (SUCCEEDED(hr))
    {
        // Tell the callback we are starting the transfer
        hr = m_spPhotoAcquireProgressCB->StartTransfer(pPhotoAcquireSource);
        if (VerifyHrSucceeded(hr))
        {
            UINT uAcquireItemCount = 0;
            hr = pPhotoAcquireSource->GetItemCount(&uAcquireItemCount);
            if (VerifyHrSucceeded(hr))
            {
                // Get the total number of items to transfer (used only for calculating progress %)
                UINT uNumItemsToTransfer = 0;
                hr = GetItemCountToTransfer(pPhotoAcquireSource, &uNumItemsToTransfer);
                if (VerifyHrSucceeded(hr))
                {
                    // Keep track of how many items out of the total to transfer we've gone through
                    // (used only for calculating progress %)
                    UINT uNumItemsTransferred = 0;
                    
                    // Loop through our list of items
                    UINT uCurrentItem = 0;
                    while (SUCCEEDED(hr) && uCurrentItem < uAcquireItemCount)
                    {
                        bool fRetryItem = false;

                        // Get the next item
                        CComPtr<IPhotoAcquireItem> spPhotoAcquireItem;
                        hr = pPhotoAcquireSource->GetItemAt(uCurrentItem, &spPhotoAcquireItem);
                        if (VerifyHrSucceeded(hr) && ShouldTransferItem(spPhotoAcquireItem))
                        {
                            // Have we been cancelled?
                            hr = AcquisitionHelpers::IsCancelled(m_spPhotoAcquireProgressCB);
                            if (SUCCEEDED(hr))
                            {
                                // Tell the callback what percentage done we are
                                hr = m_spPhotoAcquireProgressCB->UpdateTransferPercent(TRUE, MulDiv(uNumItemsTransferred, 100, uNumItemsToTransfer));
                                if (VerifyHrSucceeded(hr))
                                {
                                    // Tell the callback we are starting the transfer of this item
                                    hr = m_spPhotoAcquireProgressCB->StartItemTransfer(uCurrentItem, spPhotoAcquireItem);
                                    if (VerifyHrSucceeded(hr))
                                    {
                                        // Transfer it
                                        hr = TransferItem(pPhotoAcquireSource, spPhotoAcquireItem);

                                        // We have to tell the client we are done, even if it failed.
                                        // BUT! We don't want to overwrite an error with S_OK, so we
                                        // only save the result of EndItemTransfer if no errors had
                                        // occurred AND we get an error from the callback.
                                        HRESULT hrEnd = m_spPhotoAcquireProgressCB->EndItemTransfer(uCurrentItem, spPhotoAcquireItem, hr);
                                        if (SUCCEEDED(hr) && FAILED(hrEnd))
                                        {
                                            hr = hrEnd;
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
                                        break;

                                    case PHOTOACQUIRE_RESULT_RETRY:
                                        // Retry this item
                                        fRetryItem = true;
                                        hr = S_OK;
                                        break;

                                    case PHOTOACQUIRE_RESULT_ABORT:
                                        // Exit the loop, and set the cancelled state
                                        hr = E_ABORT;
                                        break;
                                    }
                                }
                            }

                            // Update the running count for progress.
                            if (!fRetryItem)
                            {
                                ++uNumItemsTransferred;
                            }
                        }

                        // Update the total count.
                        if (!fRetryItem)
                        {
                            ++uCurrentItem;
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
    return hr;
}

