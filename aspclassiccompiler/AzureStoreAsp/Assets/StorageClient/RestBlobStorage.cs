// ----------------------------------------------------------------------------------
// Microsoft Developer & Platform Evangelism
// 
// Copyright (c) Microsoft Corporation. All rights reserved.
// 
// THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
// EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES 
// OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
// ----------------------------------------------------------------------------------
// The example companies, organizations, products, domain names,
// e-mail addresses, logos, people, places, and events depicted
// herein are fictitious.  No association with any real company,
// organization, product, domain name, email address, logo, person,
// places, or events is intended or should be inferred.
// ----------------------------------------------------------------------------------

//
// <copyright file="RestBlobStorage.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Xml;
using Microsoft.Samples.ServiceHosting.StorageClient.StorageHttpConstants;

namespace Microsoft.Samples.ServiceHosting.StorageClient
{
    using System.Reflection;

    internal class BlobStorageRest : BlobStorage
    {
        private SharedKeyCredentials credentials;
        internal BlobStorageRest(Uri baseUri,
                                 bool? usePathStyleUris,
                                 string accountName,
                                 string base64Key
                                )
            : base(baseUri, usePathStyleUris, accountName, base64Key)
        {
            byte[] key = null;
            this.Base64Key = base64Key;
            if (base64Key != null)
                key = Convert.FromBase64String(base64Key);
            this.credentials = new SharedKeyCredentials(accountName, key);
        }

        /// <summary>
        /// Get a reference to a BlobContainer object with the given name.
        /// This method does not make any calls to the storage service.
        /// </summary>
        /// <param name="containerName">The name of the container</param>
        /// <returns>A reference to a newly created BlobContainer object</returns>
        public override BlobContainer GetBlobContainer(string containerName)
        {
            return new BlobContainerRest(
                            BaseUri,
                            UsePathStyleUris,
                            AccountName,
                            containerName,
                            Base64Key,
                            DateTime.MinValue,
                            Timeout,
                            RetryPolicy
                            );
        }

        /// <summary>
        /// Lists the containers within the account.
        /// </summary>
        /// <returns>A list of containers</returns>
        public override IEnumerable<BlobContainer> ListBlobContainers()
        {
            string marker = "", prefix = null;
            const int maxResults = StorageHttpConstants.ListingConstants.MaxContainerListResults;
            byte[] key = null;
            if (Base64Key != null)
                key = Convert.FromBase64String(Base64Key);
            SharedKeyCredentials credentials = new SharedKeyCredentials(AccountName, key);


            do
            {
                ListContainersResult result = ListContainersImpl(prefix, marker, maxResults);
                marker = result.NextMarker;
                foreach (ContainerProperties container in result.Containers)
                {
                    yield return new BlobContainerRest(
                                    BaseUri,
                                    UsePathStyleUris,
                                    AccountName,
                                    container.Name,
                                    Base64Key,
                                    container.LastModifiedTime,
                                    Timeout,
                                    RetryPolicy
                                    );
                }
            } while (marker != null);

        }

        internal class ListContainersResult
        {
            internal ListContainersResult(IEnumerable<ContainerProperties> containers, string nextMarker)
            {
                this.Containers = containers;
                this.NextMarker = nextMarker;
            }

            public IEnumerable<ContainerProperties> Containers { get; private set; }

            public string NextMarker { get; private set; }

        }

        private ListContainersResult ListContainersImpl(
            string prefix,
            string marker,
            int maxResults
            )
        {
            ListContainersResult result = null;
            ResourceUriComponents uriComponents;
            Uri listContainerssUri = CreateRequestUriForListContainers(
                                        prefix,
                                        marker,
                                        null,
                                        maxResults,
                                        out uriComponents
                                        );
            HttpWebRequest request = Utilities.CreateHttpRequest(listContainerssUri, StorageHttpConstants.HttpMethod.Get, Timeout);
            credentials.SignRequest(request, uriComponents);

            try
            {
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    if (response.StatusCode == HttpStatusCode.OK)
                    {
                        using (Stream stream = response.GetResponseStream())
                        {
                            result = ListContainersResultFromResponse(stream);
                        }
                    }
                    else
                    {
                        Utilities.ProcessUnexpectedStatusCode(response);
                    }
                }
            }
            catch (IOException ioe)
            {
                throw new StorageServerException(
                        StorageErrorCode.TransportError,
                        "The connection may be lost",
                         default(HttpStatusCode),
                         null,
                         ioe
                         );
            }
            catch (System.TimeoutException te)
            {
                throw new StorageServerException(
                              StorageErrorCode.ServiceTimeout,
                              "Timeout during listing containers",
                               HttpStatusCode.RequestTimeout,
                               null,
                               te
                               );
            }
            catch (WebException we)
            {
                throw Utilities.TranslateWebException(we);
            }
            return result;
        }

        private Uri CreateRequestUriForListContainers(
            string prefix, string fromMarker, string delimiter, int maxResults, out ResourceUriComponents uriComponents)
        {
            NameValueCollection queryParams = BlobStorageRest.CreateRequestUriForListing(prefix, fromMarker, delimiter, maxResults);

            return Utilities.CreateRequestUri(BaseUri, UsePathStyleUris, AccountName, null, null, Timeout, queryParams, out uriComponents);
        }

        internal static NameValueCollection CreateRequestUriForListing(string prefix, string fromMarker, string delimiter, int maxResults)
        {
            NameValueCollection queryParams = new NameValueCollection();
            queryParams.Add(StorageHttpConstants.QueryParams.QueryParamComp, StorageHttpConstants.CompConstants.List);

            if (!string.IsNullOrEmpty(prefix))
                queryParams.Add(StorageHttpConstants.QueryParams.QueryParamPrefix, prefix);

            if (!string.IsNullOrEmpty(fromMarker))
                queryParams.Add(StorageHttpConstants.QueryParams.QueryParamMarker, fromMarker);

            if (!string.IsNullOrEmpty(delimiter))
                queryParams.Add(StorageHttpConstants.QueryParams.QueryParamDelimiter, delimiter);

            queryParams.Add(StorageHttpConstants.QueryParams.QueryParamMaxResults,
                maxResults.ToString(CultureInfo.InvariantCulture));

            return queryParams;
        }

        private static ListContainersResult ListContainersResultFromResponse(Stream responseBody)
        {
            List<ContainerProperties> containers = new List<ContainerProperties>();
            string nextMarker = null;

            XmlDocument doc = new XmlDocument();
            try
            {
                doc.Load(responseBody);
            }
            catch (XmlException xe)
            {
                throw new StorageServerException(StorageErrorCode.ServiceBadResponse,
                    "The result of a ListBlobContainers operation could not be parsed", default(HttpStatusCode), xe);
            }


            // Get all the containers returned as the listing results
            XmlNodeList containerNodes = doc.SelectNodes(XPathQueryHelper.ContainerQuery);

            foreach (XmlNode containerNode in containerNodes)
            {
                DateTime? lastModified = XPathQueryHelper.LoadSingleChildDateTimeValue(
                    containerNode, StorageHttpConstants.XmlElementNames.LastModified, false);

                string eTag = XPathQueryHelper.LoadSingleChildStringValue(
                    containerNode, StorageHttpConstants.XmlElementNames.Etag, false);

                string containerUri = XPathQueryHelper.LoadSingleChildStringValue(
                    containerNode, StorageHttpConstants.XmlElementNames.Url, true);

                string containerName = XPathQueryHelper.LoadSingleChildStringValue(
                    containerNode, StorageHttpConstants.XmlElementNames.Name, true);

                ContainerProperties properties = new ContainerProperties(containerName);
                if (lastModified.HasValue)
                    properties.LastModifiedTime = lastModified.Value;
                properties.ETag = eTag;

                Uri uri = null;
                Uri.TryCreate(containerUri, UriKind.Absolute, out uri);
                properties.Uri = uri;

                containers.Add(properties);
            }

            // Get the nextMarker
            XmlNode nextMarkerNode = doc.SelectSingleNode(XPathQueryHelper.NextMarkerQuery);
            if (nextMarkerNode != null && nextMarkerNode.FirstChild != null)
            {
                nextMarker = nextMarkerNode.FirstChild.Value;
            }

            return new ListContainersResult(containers, nextMarker);
        }
    }

    internal class BlobContainerRest : BlobContainer
    {

        internal BlobContainerRest(
            Uri baseUri,
            bool usePathStyleUris,
            string accountName,
            string containerName,
            string base64Key,
            DateTime lastModified,
            TimeSpan timeOut,
            RetryPolicy retryPolicy
            )
            : base(baseUri, usePathStyleUris, accountName, containerName, lastModified)
        {
            ResourceUriComponents uriComponents =
                new ResourceUriComponents(accountName, containerName, null);
            containerUri = HttpRequestAccessor.ConstructResourceUri(baseUri, uriComponents, usePathStyleUris);
            if (base64Key != null)
                key = Convert.FromBase64String(base64Key);
            credentials = new SharedKeyCredentials(accountName, key);
            Timeout = timeOut;
            RetryPolicy = retryPolicy;
        }

        public override Uri ContainerUri
        {
            get
            {
                return this.containerUri;
            }
        }

        public override bool CreateContainer()
        {
            return CreateContainerImpl(null, ContainerAccessControl.Private);
        }

        /// <summary>
        /// Create the container with the specified access control if it does not exist
        /// </summary>
        /// <param name="metadata">The metadata for the container. Can be null to indicate no metadata</param>
        /// <param name="accessControl">The access control (public or private) with which to create the container</param>
        /// <returns>true if the container was created. false if the container already exists</returns>
        public override bool CreateContainer(NameValueCollection metadata, ContainerAccessControl accessControl)
        {
            return CreateContainerImpl(metadata, accessControl);
        }


        public override bool DoesContainerExist()
        {
            bool result = false;
            RetryPolicy(() =>
            {
                ResourceUriComponents uriComponents;
                Uri uri = Utilities.CreateRequestUri(BaseUri, UsePathStyleUris, AccountName, ContainerName, null, Timeout,
                            new NameValueCollection(), out uriComponents);
                HttpWebRequest request = Utilities.CreateHttpRequest(uri, StorageHttpConstants.HttpMethod.Get, Timeout);
                credentials.SignRequest(request, uriComponents);

                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode == HttpStatusCode.OK)
                            result = true;
                        else if (response.StatusCode == HttpStatusCode.Gone || response.StatusCode == HttpStatusCode.NotFound)
                            result = false;
                        else
                        {
                            Utilities.ProcessUnexpectedStatusCode(response);
                            result = false;
                        }

                        response.Close();
                    }
                }
                catch (WebException we)
                {
                    if (we.Response != null &&
                        (((HttpWebResponse)we.Response).StatusCode == HttpStatusCode.Gone ||
                         ((HttpWebResponse)we.Response).StatusCode == HttpStatusCode.NotFound)
                       )
                        result = false;
                    else
                        throw Utilities.TranslateWebException(we);
                }
            });
            return result;
        }

        /// <summary>
        /// Get the properties for the container if it exists.
        /// </summary>
        /// <returns>The metadata for the container if it exists, null otherwise</returns>
        public override ContainerProperties GetContainerProperties()
        {
            ContainerProperties result = null;
            RetryPolicy(() =>
            {
                ResourceUriComponents uriComponents;
                Uri uri = Utilities.CreateRequestUri(BaseUri, UsePathStyleUris, AccountName, ContainerName, null, Timeout,
                            new NameValueCollection(), out uriComponents);
                HttpWebRequest request = Utilities.CreateHttpRequest(uri, StorageHttpConstants.HttpMethod.Get, Timeout);
                credentials.SignRequest(request, uriComponents);

                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            result = ContainerPropertiesFromResponse(response);
                        }
                        else if (response.StatusCode == HttpStatusCode.Gone || response.StatusCode == HttpStatusCode.NotFound)
                        {
                            result = null;
                        }
                        else
                        {
                            Utilities.ProcessUnexpectedStatusCode(response);
                            result = null;
                        }
                        response.Close();
                    }
                }
                catch (WebException we)
                {
                    if (we.Response != null &&
                        ((HttpWebResponse)we.Response).StatusCode == HttpStatusCode.Gone ||
                        (((HttpWebResponse)we.Response).StatusCode == HttpStatusCode.NotFound)
                        )
                        result = null;
                    else
                        throw Utilities.TranslateWebException(we);
                }

            });
            return result;
        }

        /// <summary>
        /// Get the access control permissions associated with the container.
        /// </summary>
        /// <returns></returns>
        public override ContainerAccessControl GetContainerAccessControl()
        {
            ContainerAccessControl accessControl = ContainerAccessControl.Private;
            RetryPolicy(() =>
            {
                ResourceUriComponents uriComponents;
                NameValueCollection queryParams = new NameValueCollection();
                queryParams.Add(StorageHttpConstants.QueryParams.QueryParamComp, StorageHttpConstants.CompConstants.Acl);

                Uri uri = Utilities.CreateRequestUri(BaseUri, UsePathStyleUris, AccountName, ContainerName, null, Timeout,
                            queryParams, out uriComponents);
                HttpWebRequest request = Utilities.CreateHttpRequest(uri, StorageHttpConstants.HttpMethod.Get, Timeout);
                credentials.SignRequest(request, uriComponents);

                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            string acl = response.Headers[StorageHttpConstants.HeaderNames.PublicAccess];
                            bool publicAcl = false;
                            if (acl != null && bool.TryParse(acl, out publicAcl))
                            {
                                accessControl = (publicAcl ? ContainerAccessControl.Public : ContainerAccessControl.Private);
                            }
                            else
                            {
                                throw new StorageServerException(
                                            StorageErrorCode.ServiceBadResponse,
                                            "The server did not respond with expected container access control header",
                                            default(HttpStatusCode),
                                            null,
                                            null
                                            );
                            }
                        }
                        else
                        {
                            Utilities.ProcessUnexpectedStatusCode(response);
                        }
                        response.Close();
                    }
                }
                catch (WebException we)
                {
                    throw Utilities.TranslateWebException(we);
                }

            });
            return accessControl;
        }

        /// <summary>
        /// Get the access control permissions associated with the container.
        /// </summary>
        /// <returns></returns>
        public override void SetContainerAccessControl(ContainerAccessControl acl)
        {
            RetryPolicy(() =>
            {
                ResourceUriComponents uriComponents;
                NameValueCollection queryParams = new NameValueCollection();
                queryParams.Add(StorageHttpConstants.QueryParams.QueryParamComp, StorageHttpConstants.CompConstants.Acl);

                Uri uri = Utilities.CreateRequestUri(BaseUri, UsePathStyleUris, AccountName, ContainerName, null, Timeout,
                            queryParams, out uriComponents);
                HttpWebRequest request = Utilities.CreateHttpRequest(uri, StorageHttpConstants.HttpMethod.Put, Timeout);
                request.Headers.Add(StorageHttpConstants.HeaderNames.PublicAccess,
                    (acl == ContainerAccessControl.Public).ToString());
                credentials.SignRequest(request, uriComponents);

                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode != HttpStatusCode.OK)
                        {
                            Utilities.ProcessUnexpectedStatusCode(response);
                        }
                        response.Close();
                    }
                }
                catch (WebException we)
                {
                    throw Utilities.TranslateWebException(we);
                }

            });
        }

        public override bool DeleteContainer()
        {
            bool result = false;
            RetryPolicy(() =>
            {
                ResourceUriComponents uriComponents;
                Uri uri = Utilities.CreateRequestUri(BaseUri, this.UsePathStyleUris, AccountName, ContainerName, null, Timeout, new NameValueCollection(),
                            out uriComponents);
                HttpWebRequest request = Utilities.CreateHttpRequest(uri, StorageHttpConstants.HttpMethod.Delete, Timeout);
                credentials.SignRequest(request, uriComponents);
                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode == HttpStatusCode.OK || response.StatusCode == HttpStatusCode.Accepted)
                            result = true;
                        else
                            Utilities.ProcessUnexpectedStatusCode(response);
                        response.Close();
                    }
                }
                catch (WebException we)
                {
                    if (we.Response != null &&
                        (((HttpWebResponse)we.Response).StatusCode == HttpStatusCode.NotFound ||
                          ((HttpWebResponse)we.Response).StatusCode == HttpStatusCode.Gone ||
                          ((HttpWebResponse)we.Response).StatusCode == HttpStatusCode.Conflict))                    
                        result = false;                    
                    else                    
                        throw Utilities.TranslateWebException(we);                    
                }
            }
            );
            return result;
        }

        private bool CreateContainerImpl(NameValueCollection metadata, ContainerAccessControl accessControl)
        {
            bool result = false;
            RetryPolicy(() =>
            {
                ResourceUriComponents uriComponents;
                Uri uri = Utilities.CreateRequestUri(BaseUri, UsePathStyleUris, AccountName, ContainerName, null, Timeout, new NameValueCollection(),
                                  out uriComponents);
                HttpWebRequest request = Utilities.CreateHttpRequest(uri, StorageHttpConstants.HttpMethod.Put, Timeout);
                if (metadata != null)
                {
                    Utilities.AddMetadataHeaders(request, metadata);
                }
                if (accessControl == ContainerAccessControl.Public)
                {
                    request.Headers.Add(StorageHttpConstants.HeaderNames.PublicAccess, "true");
                }
                credentials.SignRequest(request, uriComponents);

                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode == HttpStatusCode.Created)
                            result = true;
                        else if (response.StatusCode == HttpStatusCode.Conflict)
                            result = false;
                        else
                            Utilities.ProcessUnexpectedStatusCode(response);
                        response.Close();
                    }
                }
                catch (WebException we)
                {
                    if (we.Response != null && ((HttpWebResponse)we.Response).StatusCode == HttpStatusCode.Conflict)
                        result = false;
                    else
                        throw Utilities.TranslateWebException(we);
                }
            });
            return result;
        }

        public override bool DoesBlobExist(string blobName)
        {
            //if the blob exists, the GetBlobProperties function should return for us a BlobProperties object, otherwise it returns null
            if (GetBlobProperties(blobName) == null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Create a new blob or overwrite an existing blob.
        /// </summary>
        /// 
        /// <param name="blobProperties">The properties of the blob</param>
        /// <param name="blobContents">The contents of the blob</param>
        /// <param name="overwrite">Should this request overwrite an existing blob ?</param>
        /// <returns>true if the blob was created. false if the blob already exists and <paramref name="overwrite"/>was set to false</returns>
        /// <remarks>The LastModifiedTime property of <paramref name="blobProperties"/> is set as a result of this call</remarks>
        public override bool CreateBlob(BlobProperties blobProperties, BlobContents blobContents, bool overwrite)
        {
            return PutBlobImpl(blobProperties, blobContents.AsStream, overwrite, null);
        }

        /// <summary>
        /// Updates an existing blob if it has not been modified since the specified time which is typically
        /// the last modified time of the blob when you retrieved it.
        /// Use this method to implement optimistic concurrency by avoiding clobbering changes to the blob
        /// made by another writer.
        /// </summary>
        /// <param name="blobProperties">The properties of the blob. This object should be one previously
        /// obtained from a call to GetBlob or GetBlobProperties and have its LastModifiedTime property set.</param>
        /// <param name="contents">The contents of the blob. The contents of the blob should be readable</param>
        /// <returns>true if the blob was updated. false if the blob has changed since the last time</returns>
        /// <remarks>The LastModifiedTime property of <paramref name="properties"/> is set as a result of this call</remarks>
        public override bool UpdateBlobIfNotModified(BlobProperties blobProperties, BlobContents contents)
        {
            return PutBlobImpl(blobProperties, contents.AsStream, true, blobProperties.ETag);
        }

        /// <summary>
        /// Get the blob contents and properties if the blob exisits
        /// </summary>
        /// <param name="name">The name of the blob</param>
        /// <param name="blobContents">Object in which the contents are returned.
        /// This object should contain a writable stream or should be a default constructed object.</param>
        /// <param name="transferAsChunks">Should the blob be gotten in pieces. This requires more round-trips, but will retry smaller piecs in case of failure.</param>
        /// <returns>The properties of the blob if the blob exists.</returns>
        public override BlobProperties GetBlob(string name, BlobContents blobContents, bool transferAsChunks)
        {
            bool notModified = false;
            return GetBlobImpl(name, blobContents.AsStream, null, transferAsChunks, out notModified);
        }

        /// <summary>
        /// Gets the blob contents and properties if the blob has not been modified since the time specified.
        /// Use this method if you have cached the contents of a blob and want to avoid retrieving the blob
        /// if it has not changed since the last time you retrieved it.
        /// </summary>
        /// <param name="blobProperties">The properties of the blob obtained from an earlier call to GetBlob. This
        /// parameter is updated by the call if the blob has been modified</param>
        /// <param name="blobContents">Contains the stream to which the contents of the blob are written if it has been
        /// modified</param>
        /// <param name="transferAsChunks">Should the blob be gotten in pieces. This requires more round-trips, but will retry smaller piecs in case of failure.</param>
        /// <returns>true if the blob has been modified, false otherwise</returns>
        public override bool GetBlobIfModified(BlobProperties blobProperties, BlobContents blobContents, bool transferAsChunks)
        {
            bool modified = true;
            BlobProperties newProperties =
                GetBlobImpl(blobProperties.Name, blobContents.AsStream, blobProperties.ETag, transferAsChunks, out modified);
            if (modified)
                blobProperties.Assign(newProperties);
            return modified;
        }

        /// <summary>
        /// Get the properties of the blob if it exists.
        /// This method is also the simplest way to check if a blob exists.
        /// </summary>
        /// <param name="name">The name of the blob</param>
        /// <returns>The properties of the blob if it exists. null otherwise.
        /// The properties for the contents of the blob are not set</returns>
        public override BlobProperties GetBlobProperties(string name)
        {
            try
            {
                bool modified = false;
                return GetBlobImpl(name, null, null, false, out modified);
            }
            catch (StorageClientException se)
            {
                if (se.ErrorCode == StorageErrorCode.ResourceNotFound || se.ErrorCode == StorageErrorCode.BlobNotFound)
                    return null;
                throw;
            }
        }

        /// <summary>
        /// Set the metadata of an existing blob.
        /// </summary>
        /// <param name="blobProperties">The blob properties object whose metadata is to be updated</param>
        public override void UpdateBlobMetadata(BlobProperties blobProperties)
        {
            SetBlobMetadataImpl(blobProperties, null);
        }

        /// <summary>
        /// Set the metadata of an existing blob if it has not been modified since it was last retrieved.
        /// </summary>
        /// <param name="blobProperties">The blob properties object whose metadata is to be updated.
        /// Typically obtained by a previous call to GetBlob or GetBlobProperties</param>
        public override bool UpdateBlobMetadataIfNotModified(BlobProperties blobProperties)
        {
            return SetBlobMetadataImpl(blobProperties, blobProperties.ETag);
        }

        /// <summary>
        /// Delete a blob with the given name
        /// </summary>
        /// <param name="name">The name of the blob</param>
        /// <returns>true if the blob exists and was successfully deleted, false if the blob does not exist</returns>
        public override bool DeleteBlob(string name)
        {
            bool unused = false;
            return DeleteBlobImpl(name, null, out unused);
        }

        /// <summary>
        /// Delete a blob with the given name if the blob has not been modified since it was last obtained.
        /// Use this method for optimistic concurrency to avoid deleting a blob that has been modified since
        /// the last time you retrieved it
        /// </summary>
        /// <param name="blob">A blob object (typically previously obtained from a GetBlob call)</param>
        /// <param name="modified">This out parameter is set to true if the blob was not deleted because
        /// it was modified</param>
        /// <returns>true if the blob exists and was successfully deleted, false if the blob does not exist or was
        /// not deleted because the blob was modified.</returns>
        public override bool DeleteBlobIfNotModified(BlobProperties blob, out bool modified)
        {
            return DeleteBlobImpl(blob.Name, blob.ETag, out modified);
        }

        /// <summary>
        /// Enumerates all blobs with a given prefix.
        /// </summary>
        /// <param name="prefix"></param>
        /// <param name="combineCommonPrefixes">If true common prefixes with "/" as seperator</param>
        /// <returns>The list of blob properties and common prefixes</returns>
        public override IEnumerable<object> ListBlobs(string prefix, bool combineCommonPrefixes)
        {
            string marker = "";
            const int maxResults = StorageHttpConstants.ListingConstants.MaxBlobListResults;

            string delimiter = combineCommonPrefixes ? "/" : "";
            do
            {
                ListBlobsResult result = ListBlobsImpl(prefix, marker, delimiter, maxResults);
                marker = result.NextMarker;
                foreach (string commonPrefix in result.CommonPrefixes)
                {
                    yield return commonPrefix;
                }

                foreach (BlobProperties blob in result.Blobs)
                {
                    yield return blob;
                }
            } while (marker != null);
        }

        private HttpWebRequest CreateHttpRequestForPutBlob(Uri uri, string httpMethod, BlobProperties blobProperties,
                    long contentLength, bool overwrite, string eTag)
        {
            HttpWebRequest request = Utilities.CreateHttpRequest(uri, httpMethod, Timeout);
            if (blobProperties.ContentEncoding != null)
                request.Headers.Add(StorageHttpConstants.HeaderNames.ContentEncoding, blobProperties.ContentEncoding);
            if (blobProperties.ContentLanguage != null)
                request.Headers.Add(StorageHttpConstants.HeaderNames.ContentLanguage, blobProperties.ContentLanguage);
            if (blobProperties.ContentType != null)
                request.ContentType = blobProperties.ContentType;
            if (eTag != null)
                request.Headers.Add(StorageHttpConstants.HeaderNames.IfMatch, eTag);

            if (blobProperties.Metadata != null && blobProperties.Metadata.Count > 0)
            {
                Utilities.AddMetadataHeaders(request, blobProperties.Metadata);
            }
            request.ContentLength = contentLength;
            if (!overwrite)
            {
                request.Headers.Set(StorageHttpConstants.HeaderNames.IfNoneMatch, "*");
            }
            return request;
        }

        private HttpWebRequest CreateHttpRequestForGetBlob(Uri uri, string httpMethod, string ifNoneMatchETag, string ifMatchETag)
        {
            HttpWebRequest request = Utilities.CreateHttpRequest(uri, httpMethod, Timeout);
            if (ifNoneMatchETag != null)
            {
                request.Headers.Add(StorageHttpConstants.HeaderNames.IfNoneMatch, ifNoneMatchETag);
            }
            if (ifMatchETag != null)
            {
                request.Headers.Add(StorageHttpConstants.HeaderNames.IfMatch, ifMatchETag);
            }
            return request;
        }


        /// <summary>
        /// Uploads a blob in chunks.
        /// </summary>
        /// <param name="blobProperties"></param>
        /// <param name="stream"></param>
        /// <param name="overwrite"></param>
        /// <param name="eTag"></param>
        /// <returns></returns>
        private bool PutLargeBlobImpl(BlobProperties blobProperties, Stream stream, bool overwrite, string eTag)
        {
            bool retval = false;
            // Since we got a large block, chunk it into smaller pieces called blocks
            long blockSize = StorageHttpConstants.BlobBlockConstants.BlockSize;
            long startPosition = stream.Position;
            long length = stream.Length - startPosition;
            int numBlocks = (int)Math.Ceiling((double)length / blockSize);
            string[] blockIds = new string[numBlocks];

            //We can retry only if the stream supports seeking. An alternative is to buffer the data in memory
            //but we do not do this currently.
            RetryPolicy R = (stream.CanSeek ? this.RetryPolicy : RetryPolicies.NoRetry);

            //Upload each of the blocks, retrying any failed uploads
            for (int i = 0; i < numBlocks; ++i)
            {
                string blockId = Convert.ToBase64String(System.BitConverter.GetBytes(i));
                blockIds[i] = blockId;
                R(() =>
                {
                    // Rewind the stream to appropriate location in case this is a retry
                    if (stream.CanSeek)
                        stream.Position = startPosition + i * blockSize;
                    NameValueCollection nvc = new NameValueCollection();
                    nvc.Add(QueryParams.QueryParamComp, CompConstants.Block);
                    nvc.Add(QueryParams.QueryParamBlockId, blockId); // The block naming should be more elaborate to give more meanings on GetBlockList
                    long blockLength = Math.Min(blockSize, length - stream.Position);
                    retval = UploadData(blobProperties, stream, blockLength, overwrite, eTag, nvc);
                });
            }

            // Now commit the list
            // First create the output
            using (MemoryStream buffer = new MemoryStream())
            {
                // construct our own XML segment with correct blockId's
                XmlTextWriter writer = new XmlTextWriter(buffer, Encoding.UTF8);
                writer.WriteStartDocument();
                writer.WriteStartElement(XmlElementNames.BlockList);
                foreach (string id in blockIds)
                {
                    writer.WriteElementString(XmlElementNames.Block, id);
                }
                writer.WriteEndElement();
                writer.WriteEndDocument();
                writer.Flush();
                buffer.Position = 0; //Rewind

                NameValueCollection nvc = new NameValueCollection();
                nvc.Add(QueryParams.QueryParamComp, CompConstants.BlockList);

                retval = UploadData(blobProperties, buffer, buffer.Length, overwrite, eTag, nvc);
            }

            return retval;
        }

        private bool PutBlobImpl(BlobProperties blobProperties, Stream stream, bool overwrite, string eTag)
        {
            // If the blob is large, we should use blocks to upload it in pieces.
            // This will ensure that a broken connection will only impact a single piece
            long originalPosition = stream.Position;
            long length = stream.Length - stream.Position;
            if (length > StorageHttpConstants.BlobBlockConstants.MaximumBlobSizeBeforeTransmittingAsBlocks)
                return PutLargeBlobImpl(blobProperties, stream, overwrite, eTag);

            bool retval = false;
            RetryPolicy R = stream.CanSeek ? this.RetryPolicy : RetryPolicies.NoRetry;
            R(() =>
            {
                if (stream.CanSeek)
                    stream.Position = originalPosition;
                retval = UploadData(blobProperties, stream, length, overwrite, eTag, new NameValueCollection());
            });

            return retval;
        }

        private bool UploadData(BlobProperties blobProperties, Stream stream, long length, bool overwrite, string eTag, NameValueCollection nvc)
        {
            ResourceUriComponents uriComponents;
            Uri blobUri = Utilities.CreateRequestUri(BaseUri, this.UsePathStyleUris, AccountName, ContainerName,
                 blobProperties.Name, Timeout, nvc, out uriComponents);
            HttpWebRequest request = CreateHttpRequestForPutBlob(
                                        blobUri,
                                        StorageHttpConstants.HttpMethod.Put,
                                        blobProperties,
                                        length,
                                        overwrite,
                                        eTag
                                        );
            credentials.SignRequest(request, uriComponents);
            bool retval = false;

            try
            {
                using (Stream requestStream = request.GetRequestStream())
                {
                    Utilities.CopyStream(stream, requestStream, length);
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode == HttpStatusCode.Created)
                        {
                            retval = true;
                        }
                        else if (!overwrite && 
                                (response.StatusCode == HttpStatusCode.PreconditionFailed ||
                                 response.StatusCode == HttpStatusCode.NotModified))

                        {
                            retval = false;
                        }
                        else
                        {
                            retval = false;
                            Utilities.ProcessUnexpectedStatusCode(response);
                        }

                        blobProperties.LastModifiedTime = response.LastModified.ToUniversalTime();
                        blobProperties.ETag = response.Headers[StorageHttpConstants.HeaderNames.ETag];
                        response.Close();
                    }
                    requestStream.Close();
                }
            }
            catch (IOException ioe)
            {
                throw new StorageServerException(
                            StorageErrorCode.TransportError,
                            "The connection may be lost",
                            default(HttpStatusCode),
                            null,
                            ioe
                            );
            }
            catch (System.TimeoutException te)
            {
                throw new StorageServerException(
                            StorageErrorCode.ServiceTimeout,
                            "Timeout during blob upload",
                            HttpStatusCode.RequestTimeout,
                            null,
                            te
                            );
            }
            catch (WebException we)
            {
                if (we.Response != null)
                {
                    HttpWebResponse response = (HttpWebResponse)we.Response;


                    if ((response.StatusCode == HttpStatusCode.PreconditionFailed ||
                         response.StatusCode == HttpStatusCode.NotModified) &&
                        (!overwrite || eTag != null))

                    {
                        retval = false;
                        return retval;
                    }
                }
                throw Utilities.TranslateWebException(we);
            }
            return retval;
        }

        private BlobProperties GetBlobImpl(string blobName, Stream stream, string oldETag, bool chunked, out bool modified)
        {
            //We introduce this local variable since lambda expressions cannot contain use an out parameter
            BlobProperties blobProperties = null;
            bool localModified = true;

            //If we are interested only in the blob properties (stream == null) or we are performing
            // a chunked download we first obtain just the blob properties
            if (stream == null || chunked)
            {
                RetryPolicy(() =>
                {
                    blobProperties = DownloadData(
                                        blobName,
                                        null,
                                        oldETag,
                                        null,
                                        0,
                                        0,
                                        new NameValueCollection(),
                                        ref localModified
                                        );
                });
                modified = localModified;
                if (stream == null)
                {
                    return blobProperties;
                }
            }


            RetryPolicy R = stream.CanSeek ? this.RetryPolicy : RetryPolicies.NoRetry;
            long originalPosition = stream.CanSeek ? stream.Position : 0;
            if (chunked && blobProperties != null && blobProperties.ContentLength > 0)
            {
                //Chunked download. Obtain ranges of the blobs in 'BlockSize' chunks
                //Ensure that the If-Match <Etag>header is used on each request so
                //that we are assured that all data belongs to the single blob we
                //started downloading.
                long location = 0;
                while (location < blobProperties.ContentLength)
                {
                    long nBytes = Math.Min(blobProperties.ContentLength - location, StorageHttpConstants.BlobBlockConstants.BlockSize);
                    R(() =>
                    {
                        // Set the position to rewind in case of a retry
                        if (stream.CanSeek)
                            stream.Position = originalPosition + location;
                        DownloadData(
                                blobName,
                                stream,
                                oldETag,
                                blobProperties.ETag,
                                location,
                                nBytes,
                                new NameValueCollection(),
                                ref localModified
                                );
                    });
                    location += nBytes;
                }
            }
            else
            {
                //Non-chunked download. Obtain the entire blob in a single request.
                R(() =>
                {
                    // Set the position to rewind in case of a retry
                    if (stream.CanSeek)
                        stream.Position = originalPosition;

                    blobProperties = DownloadData(
                                        blobName,
                                        stream,
                                        oldETag,
                                        null,
                                        0,
                                        0,
                                        new NameValueCollection(),
                                        ref localModified
                                        );
                });
            }
            modified = localModified;
            return blobProperties;
        }

        private bool SetBlobMetadataImpl(BlobProperties blobProperties, string eTag)
        {
            bool retval = false;
            RetryPolicy(() =>
            {
                NameValueCollection queryParams = new NameValueCollection();
                queryParams.Add(StorageHttpConstants.QueryParams.QueryParamComp, StorageHttpConstants.CompConstants.Metadata);

                ResourceUriComponents uriComponents;
                Uri uri = Utilities.CreateRequestUri(
                    BaseUri,
                    this.UsePathStyleUris,
                    AccountName,
                    ContainerName,
                    blobProperties.Name,
                    Timeout,
                    queryParams,
                    out uriComponents
                    );
                HttpWebRequest request = Utilities.CreateHttpRequest(uri, StorageHttpConstants.HttpMethod.Put, Timeout);
                if (blobProperties.Metadata != null)
                {
                    Utilities.AddMetadataHeaders(request, blobProperties.Metadata);
                }
                if (eTag != null)
                {
                    request.Headers.Add(StorageHttpConstants.HeaderNames.IfMatch, eTag);
                }
                credentials.SignRequest(request, uriComponents);

                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            retval = true;
                        }
                        else if ((response.StatusCode == HttpStatusCode.PreconditionFailed ||
                                  response.StatusCode == HttpStatusCode.NotModified) && 
                                 eTag != null)
                        {
                            retval = false;
                        }
                        else
                        {
                            Utilities.ProcessUnexpectedStatusCode(response);
                            retval = false;
                        }
                        blobProperties.LastModifiedTime = response.LastModified.ToUniversalTime();
                        blobProperties.ETag = response.Headers[StorageHttpConstants.HeaderNames.ETag];
                        response.Close();
                    }
                }
                catch (IOException ioe)
                {
                    throw new StorageServerException(StorageErrorCode.TransportError, "The connection may be lost",
                                default(HttpStatusCode), ioe);
                }
                catch (System.TimeoutException te)
                {
                    throw new StorageServerException(StorageErrorCode.ServiceTimeout, "Timeout during blob metadata upload",
                                    HttpStatusCode.RequestTimeout, te);
                }
                catch (WebException we)
                {
                    if (we.Response != null)
                    {
                        HttpWebResponse response = (HttpWebResponse)we.Response;
                        if (eTag != null &&
                            (response.StatusCode == HttpStatusCode.PreconditionFailed ||
                             response.StatusCode == HttpStatusCode.NotModified))
                        {
                            retval = false;
                            return;
                        }
                    }
                    throw Utilities.TranslateWebException(we);
                }
            });
            return retval;
        }

        private List<int> GetBlockList(string blobName, string eTag)
        {
            List<int> blocks;
            using (MemoryStream blockListStream = new MemoryStream())
            {
                bool temp = true;
                NameValueCollection nvc = new NameValueCollection();
                nvc.Add(QueryParams.QueryParamComp, CompConstants.BlockList);

                DownloadData(blobName, blockListStream, eTag, null, 0, 0, nvc, ref temp);
                XmlDocument doc = new XmlDocument();
                blockListStream.Position = 0;
                doc.Load(blockListStream);

                blocks = new List<int>();
                foreach (XmlNode block in doc.SelectNodes(XPathQueryHelper.BlockQuery))
                {
                    blocks.Add((int)XPathQueryHelper.LoadSingleChildLongValue(block, XmlElementNames.Size, false));
                }
            }
            return blocks;
        }

        /// <summary>
        /// Helper method used for getting blobs, ranges of blobs and blob properties.
        /// </summary>
        /// <param name="blobName">Name of the blob</param>
        /// <param name="stream">The output stream to write blob data to. Can be null if only retrieving blob properties</param>
        /// <param name="eTagIfNoneMatch">The If-None-Match header. Used to avoid downloading blob data if the blob has not changed</param>
        /// <param name="eTagIfMatch">The If-Match header. Used to ensure that all chunks of the blob are of the same blob</param>
        /// <param name="offset">The offset of the blob data to begin downloading from. Set to 0 to download all data.</param>
        /// <param name="length">The length of the blob data to download. Set to 0 to download all data</param>
        /// <param name="nvc">Query paramters to add to the request.</param>
        /// <param name="localModified">Whether the blob had been modfied with respect to the <paramref name="eTagIfNoneMatch"/></param>
        /// <returns></returns>
        private BlobProperties DownloadData(
            string blobName,
            Stream stream,
            string eTagIfNoneMatch,
            string eTagIfMatch,
            long offset,
            long length,
            NameValueCollection nvc,
            ref bool localModified
            )
        {
            ResourceUriComponents uriComponents;
            Uri blobUri = Utilities.CreateRequestUri(BaseUri, this.UsePathStyleUris, AccountName, ContainerName, blobName, Timeout, nvc, out uriComponents);
            string httpMethod = (stream == null ? StorageHttpConstants.HttpMethod.Head : StorageHttpConstants.HttpMethod.Get);
            HttpWebRequest request = CreateHttpRequestForGetBlob(blobUri, httpMethod, eTagIfNoneMatch, eTagIfMatch);

            if (offset != 0 || length != 0)
            {
                //Use the blob storage custom header for range since the standard HttpWebRequest.AddRange 
                //accepts only 32 bit integers and so does not work for large blobs
                string rangeHeaderValue = string.Format(
                                            CultureInfo.InvariantCulture,
                                            HeaderValues.RangeHeaderFormat,
                                            offset,
                                            offset + length - 1);
                request.Headers.Add(HeaderNames.StorageRange, rangeHeaderValue);
            }
            credentials.SignRequest(request, uriComponents);

            BlobProperties blobProperties;

            try
            {
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    if (response.StatusCode == HttpStatusCode.OK
                        || response.StatusCode == HttpStatusCode.PartialContent)
                    {
                        blobProperties = BlobPropertiesFromResponse(blobName, blobUri, response);
                        if (stream != null)
                        {
                            using (Stream responseStream = response.GetResponseStream())
                            {
                                long bytesCopied = Utilities.CopyStream(responseStream, stream);
                                if (blobProperties.ContentLength > 0 && bytesCopied < blobProperties.ContentLength)
                                    throw new StorageServerException(
                                        StorageErrorCode.ServiceTimeout,
                                        "Unable to read complete data from server",
                                        HttpStatusCode.RequestTimeout,
                                        null
                                        );
                            }
                        }
                    }
                    else
                    {
                        Utilities.ProcessUnexpectedStatusCode(response);
                        return null;
                    }
                }
                return blobProperties;
            }
            catch (IOException ioe)
            {
                throw new StorageServerException(StorageErrorCode.TransportError, "The connection may be lost",
                            default(HttpStatusCode), ioe);
            }
            catch (System.TimeoutException te)
            {
                throw new StorageServerException(StorageErrorCode.ServiceTimeout, "Timeout during blob download",
                                HttpStatusCode.RequestTimeout, te);
            }
            catch (WebException we)
            {
                if (we.Response != null)
                {
                    HttpWebResponse response = (HttpWebResponse)we.Response;
                    if (eTagIfNoneMatch != null &&
                       (response.StatusCode == HttpStatusCode.PreconditionFailed ||
                         response.StatusCode == HttpStatusCode.NotModified))
                    {
                        localModified = false;
                        blobProperties = null;
                        return blobProperties;
                    }
                }
                throw Utilities.TranslateWebException(we);
            }
        }

        private static NameValueCollection MetadataFromHeaders(WebHeaderCollection headers)
        {
            int prefixLength = StorageHttpConstants.HeaderNames.PrefixForMetadata.Length;
            string[] headerNames = headers.AllKeys;
            NameValueCollection metadataEntries = new NameValueCollection();
            foreach (string headerName in headerNames)
            {
                if (headerName.StartsWith(StorageHttpConstants.HeaderNames.PrefixForMetadata,
                        StringComparison.OrdinalIgnoreCase))
                {
                    // strip out the metadata prefix
                    metadataEntries.Add(headerName.Substring(prefixLength), headers[headerName]);
                }
            }
            return metadataEntries;
        }

        private static BlobProperties BlobPropertiesFromResponse(
                            string blobName, Uri blobUri, HttpWebResponse response
                            )
        {
            BlobProperties blobProperties = new BlobProperties(blobName);
            blobProperties.Uri = blobUri;
            blobProperties.ContentEncoding = response.Headers[StorageHttpConstants.HeaderNames.ContentEncoding];
            blobProperties.LastModifiedTime = response.LastModified.ToUniversalTime();
            blobProperties.ETag = response.Headers[StorageHttpConstants.HeaderNames.ETag];
            blobProperties.ContentLanguage = response.Headers[StorageHttpConstants.HeaderNames.ContentLanguage];
            blobProperties.ContentLength = response.ContentLength;
            blobProperties.ContentType = response.ContentType;

            NameValueCollection metadataEntries = MetadataFromHeaders(response.Headers);
            if (metadataEntries.Count > 0)
                blobProperties.Metadata = metadataEntries;

            return blobProperties;
        }

        private ContainerProperties ContainerPropertiesFromResponse(HttpWebResponse response)
        {
            ContainerProperties prop = new ContainerProperties(ContainerName);
            prop.LastModifiedTime = response.LastModified.ToUniversalTime();
            prop.ETag = response.Headers[StorageHttpConstants.HeaderNames.ETag];
            prop.Uri = containerUri;
            prop.Metadata = MetadataFromHeaders(response.Headers);
            return prop;
        }

        bool DeleteBlobImpl(string name, string eTag, out bool modified)
        {
            bool retval = false;
            bool localModified = false;
            RetryPolicy(() =>
            {
                ResourceUriComponents uriComponents;
                Uri blobUri = Utilities.CreateRequestUri(BaseUri, this.UsePathStyleUris, AccountName, ContainerName, name, Timeout, new NameValueCollection(),
                                           out uriComponents);
                HttpWebRequest request = Utilities.CreateHttpRequest(blobUri, StorageHttpConstants.HttpMethod.Delete, Timeout);

                if (eTag != null)
                    request.Headers.Add(StorageHttpConstants.HeaderNames.IfMatch, eTag);
                credentials.SignRequest(request, uriComponents);

                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode == HttpStatusCode.OK ||
                            response.StatusCode == HttpStatusCode.Accepted)
                        {
                            retval = true;
                        }
                        else
                            Utilities.ProcessUnexpectedStatusCode(response);
                        response.Close();
                    }
                }
                catch (IOException ioe)
                {
                    throw new StorageServerException(StorageErrorCode.TransportError, "The connection may be lost",
                                default(HttpStatusCode), ioe);
                }
                catch (System.TimeoutException te)
                {
                    throw new StorageServerException(StorageErrorCode.ServiceTimeout, "Timeout during blob delete",
                                    HttpStatusCode.RequestTimeout, te);
                }
                catch (WebException we)
                {
                    if (we.Response != null)
                    {
                        HttpStatusCode status = ((HttpWebResponse)we.Response).StatusCode;
                        if (status == HttpStatusCode.NotFound ||
                            status == HttpStatusCode.Gone)
                        {
                            localModified = true;
                            retval = false;
                        }
                        else if (status == HttpStatusCode.PreconditionFailed ||
                                 status == HttpStatusCode.NotModified)
                        {
                            retval = false;
                            localModified = true;
                        }
                        else
                        {
                            throw Utilities.TranslateWebException(we);
                        }
                    }
                    else
                    {
                        throw Utilities.TranslateWebException(we);
                    }
                }
            });
            modified = localModified;
            return retval;
        }

        internal class ListBlobsResult
        {
            internal ListBlobsResult(IEnumerable<BlobProperties> blobs, IEnumerable<string> commonPrefixes, string nextMarker)
            {
                Blobs = blobs;
                CommonPrefixes = commonPrefixes;
                NextMarker = nextMarker;
            }

            internal IEnumerable<BlobProperties> Blobs
            {
                get;
                private set;
            }

            internal IEnumerable<string> CommonPrefixes
            {
                get;
                private set;
            }

            internal string NextMarker
            {
                get;
                private set;
            }
        }

        private ListBlobsResult ListBlobsImpl(string prefix, string fromMarker, string delimiter, int maxCount)
        {
            ListBlobsResult result = null;

            RetryPolicy(() =>
            {
                ResourceUriComponents uriComponents;
                Uri listBlobsUri = CreateRequestUriForListBlobs(prefix, fromMarker, delimiter, maxCount, out uriComponents);

                HttpWebRequest request = Utilities.CreateHttpRequest(listBlobsUri, StorageHttpConstants.HttpMethod.Get, Timeout);
                credentials.SignRequest(request, uriComponents);


                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            using (Stream stream = response.GetResponseStream())
                            {
                                result = ListBlobsResultFromResponse(stream);
                            }
                        }
                        else
                        {
                            Utilities.ProcessUnexpectedStatusCode(response);
                        }
                    }
                }
                catch (IOException ioe)
                {
                    throw new StorageServerException(StorageErrorCode.TransportError, "The connection may be lost",
                                default(HttpStatusCode), ioe);
                }
                catch (System.TimeoutException te)
                {
                    throw new StorageServerException(StorageErrorCode.ServiceTimeout, "Timeout during listing blobs",
                                    HttpStatusCode.RequestTimeout, te);
                }
                catch (WebException we)
                {
                    throw Utilities.TranslateWebException(we);
                }
            });
            return result;
        }

        private Uri CreateRequestUriForListBlobs(
            string prefix, string fromMarker, string delimiter, int maxResults, out ResourceUriComponents uriComponents)
        {
            NameValueCollection queryParams = BlobStorageRest.CreateRequestUriForListing(prefix, fromMarker, delimiter, maxResults);
            return Utilities.CreateRequestUri(BaseUri, this.UsePathStyleUris, AccountName, ContainerName, null, Timeout, queryParams, out uriComponents);
        }

        private static ListBlobsResult ListBlobsResultFromResponse(Stream responseBody)
        {
            List<BlobProperties> blobs = new List<BlobProperties>();
            List<string> commonPrefixes = new List<string>();
            string nextMarker = null;

            XmlDocument doc = new XmlDocument();
            try
            {
                doc.Load(responseBody);
            }
            catch (XmlException xe)
            {
                throw new StorageServerException(StorageErrorCode.ServiceBadResponse,
                    "The result of a ListBlobs operation could not be parsed", default(HttpStatusCode), xe);
            }

            // Get the commonPrefixes
            XmlNodeList prefixNodes = doc.SelectNodes(XPathQueryHelper.CommonPrefixQuery);

            foreach (XmlNode prefixNode in prefixNodes)
            {
                string blobPrefix = XPathQueryHelper.LoadSingleChildStringValue(
                    prefixNode, StorageHttpConstants.XmlElementNames.BlobPrefixName, false);
                commonPrefixes.Add(blobPrefix);
            }

            // Get all the blobs returned as the listing results
            XmlNodeList blobNodes = doc.SelectNodes(XPathQueryHelper.BlobQuery);

            foreach (XmlNode blobNode in blobNodes)
            {
                DateTime? lastModified = XPathQueryHelper.LoadSingleChildDateTimeValue(
                    blobNode, StorageHttpConstants.XmlElementNames.LastModified, false);

                string eTag = XPathQueryHelper.LoadSingleChildStringValue(
                    blobNode, StorageHttpConstants.XmlElementNames.Etag, false);

                string contentType = XPathQueryHelper.LoadSingleChildStringValue(
                    blobNode, StorageHttpConstants.XmlElementNames.ContentType, false);

                string contentEncoding = XPathQueryHelper.LoadSingleChildStringValue(
                    blobNode, StorageHttpConstants.XmlElementNames.ContentEncoding, false);

                string contentLanguage = XPathQueryHelper.LoadSingleChildStringValue(
                    blobNode, StorageHttpConstants.XmlElementNames.ContentLanguage, false);

                long? blobSize = XPathQueryHelper.LoadSingleChildLongValue(
                    blobNode, StorageHttpConstants.XmlElementNames.Size, false);

                string blobUrl = XPathQueryHelper.LoadSingleChildStringValue(
                    blobNode, StorageHttpConstants.XmlElementNames.Url, true);

                string blobName = XPathQueryHelper.LoadSingleChildStringValue(
                    blobNode, StorageHttpConstants.XmlElementNames.BlobName, true);

                BlobProperties properties = new BlobProperties(blobName);
                properties.Uri = new Uri(blobUrl);
                if (lastModified.HasValue)
                    properties.LastModifiedTime = lastModified.Value;
                properties.ContentEncoding = contentEncoding;
                properties.ContentLanguage = contentLanguage;
                properties.ETag = eTag;
                properties.ContentLength = (blobSize == null ? 0 : blobSize.Value);
                properties.ContentType = contentType;

                blobs.Add(properties);
            }

            // Get the nextMarker
            XmlNode nextMarkerNode = doc.SelectSingleNode(XPathQueryHelper.NextMarkerQuery);
            if (nextMarkerNode != null && nextMarkerNode.FirstChild != null)
            {
                nextMarker = nextMarkerNode.FirstChild.Value;
            }

            return new ListBlobsResult(blobs, commonPrefixes, nextMarker);
        }

        private Uri containerUri;
        private byte[] key;
        private SharedKeyCredentials credentials;
    }

    /// <summary>
    /// Helper class for loading values from an XML segment
    /// </summary>
    internal static class XPathQueryHelper
    {
        // In general queries are like "//{0}/{1}/{2}" - using Join as it's more efficient than Format

        internal static readonly string NextMarkerQuery = string.Join(StorageHttpConstants.ConstChars.Slash,
            new string[] 
            {
                "", "", 
                StorageHttpConstants.XmlElementNames.EnumerationResults,
                StorageHttpConstants.XmlElementNames.NextMarker
            });

        internal static readonly string ContainerQuery = string.Join(StorageHttpConstants.ConstChars.Slash,
            new string[] 
            {
                "", "", 
                StorageHttpConstants.XmlElementNames.EnumerationResults,
                StorageHttpConstants.XmlElementNames.Containers,
                StorageHttpConstants.XmlElementNames.Container
            });

        internal static readonly string BlobQuery = string.Join(StorageHttpConstants.ConstChars.Slash,
            new string[] 
            {
                "", "", 
                StorageHttpConstants.XmlElementNames.EnumerationResults,
                StorageHttpConstants.XmlElementNames.Blobs,
                StorageHttpConstants.XmlElementNames.Blob
            });

        internal static readonly string BlockQuery = string.Join(StorageHttpConstants.ConstChars.Slash,
            new string[] 
            {
                "", "", 
                StorageHttpConstants.XmlElementNames.BlockList,
                StorageHttpConstants.XmlElementNames.Block
            });

        internal static readonly string QueueListQuery = string.Join(StorageHttpConstants.ConstChars.Slash,
            new string[] 
            {
                "", "", 
                StorageHttpConstants.XmlElementNames.EnumerationResults,
                StorageHttpConstants.XmlElementNames.Queues,
                StorageHttpConstants.XmlElementNames.Queue
            });

        internal static readonly string MessagesListQuery = string.Join(StorageHttpConstants.ConstChars.Slash,
                    new string[] 
            {
                "", "", 
                StorageHttpConstants.XmlElementNames.QueueMessagesList,
                StorageHttpConstants.XmlElementNames.QueueMessage,
            });

        internal static readonly string CommonPrefixQuery = string.Join(StorageHttpConstants.ConstChars.Slash,
            new string[] 
            {
                "", "", 
                StorageHttpConstants.XmlElementNames.EnumerationResults,
                StorageHttpConstants.XmlElementNames.Blobs,
                StorageHttpConstants.XmlElementNames.BlobPrefix
            });

        internal static DateTime? LoadSingleChildDateTimeValue(XmlNode node, string childName, bool throwIfNotFound)
        {
            XmlNode childNode = node.SelectSingleNode(childName);

            if (childNode != null && childNode.FirstChild != null)
            {
                DateTime? dateTime;
                if (!Utilities.TryGetDateTimeFromHttpString(childNode.FirstChild.Value, out dateTime))
                {
                    throw new StorageServerException(StorageErrorCode.ServiceBadResponse,
                        "Date time value returned from server " + childNode.FirstChild.Value + " can't be parsed.",
                        default(HttpStatusCode),
                        null
                        );
                }
                return dateTime;
            }
            else if (!throwIfNotFound)
            {
                return null;
            }
            else
            {
                return null;
            }
        }


        internal static string LoadSingleChildStringValue(XmlNode node, string childName, bool throwIfNotFound)
        {
            XmlNode childNode = node.SelectSingleNode(childName);

            if (childNode != null && childNode.FirstChild != null)
            {
                return childNode.FirstChild.Value;
            }
            else if (!throwIfNotFound)
            {
                return null;
            }
            else
            {
                return null;   // unnecessary since Fail will throw, but keeps the compiler happy
            }
        }

        internal static long? LoadSingleChildLongValue(XmlNode node, string childName, bool throwIfNotFound)
        {
            XmlNode childNode = node.SelectSingleNode(childName);

            if (childNode != null && childNode.FirstChild != null)
            {
                return long.Parse(childNode.FirstChild.Value, CultureInfo.InvariantCulture);
            }
            else if (!throwIfNotFound)
            {
                return null;
            }
            else
            {
                return null;   // unnecessary since Fail will throw, but keeps the compiler happy
            }
        }
    }
}