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
// <copyright file="BlobStorage.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//
using System;
using System.Collections.Generic;
using System.IO;
using System.Collections.Specialized;
using System.Threading;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Diagnostics;

[assembly:CLSCompliant(true)]

// disable the generation of warnings for missing documentation elements for 
// public classes/members in this file
#pragma warning disable 1591

namespace Microsoft.Samples.ServiceHosting.StorageClient
{

    /// <summary>
    /// This delegate define the shape of a retry policy. A retry policy will invoke the given
    /// <paramref name="action"/> as many times as it wants to in the face of 
    /// retriable StorageServerExceptions.
    /// </summary>
    /// <param name="action">The action to retry</param>
    /// <returns></returns>
    public delegate void RetryPolicy(Action action);

    #region Blob Storage API
    /// <summary>
    /// The entry point of the blob storage API
    /// </summary>
    public abstract class BlobStorage
    {

        /// <summary>
        /// Factory method for BlobStorage
        /// </summary>
        /// <param name="baseUri">The base URI of the blob storage service</param>
        /// <param name="usePathStyleUris">If true, path-style URIs (http://baseuri/accountname/containername/objectname) are used.
        /// If false host-style URIs (http://accountname.baseuri/containername/objectname) are used,
        /// where baseuri is the URI of the service.
        /// If null, the choice is made automatically: path-style URIs if host name part of base URI is an 
        /// IP addres, host-style otherwise.</param>
        /// <param name="accountName">The name of the storage account</param>
        /// <param name="base64Key">Authentication key used for signing requests</param>
        /// <returns>A newly created BlobStorage instance</returns>
        public static BlobStorage Create(
                                    Uri baseUri,
                                    bool? usePathStyleUris,
                                    string accountName,
                                    string base64Key
                                    )
        {
            //We create a StorageAccountInfo and then extract the properties of that object.
            //This is because the constructor of StorageAccountInfo does normalization of BaseUri.
            StorageAccountInfo accountInfo = new StorageAccountInfo(
                                                baseUri,
                                                usePathStyleUris,
                                                accountName,
                                                base64Key
                                                );
            return new BlobStorageRest(
                accountInfo.BaseUri,
                accountInfo.UsePathStyleUris,
                accountInfo.AccountName,
                accountInfo.Base64Key
                );
        }

        /// <summary>
        /// Factory method for BlobStorage
        /// </summary>
        /// <param name="accountInfo">Account information</param>
        /// <returns>A newly created BlobStorage instance</returns>
        public static BlobStorage Create(StorageAccountInfo accountInfo)
        {
            return new BlobStorageRest(
                accountInfo.BaseUri,
                accountInfo.UsePathStyleUris,
                accountInfo.AccountName,
                accountInfo.Base64Key
                );
        }
                                    

        /// <summary>
        /// Get a reference to a newly created BlobContainer object.
        /// This method does not make any calls to the storage service.
        /// </summary>
        /// <param name="containerName">The name of the container</param>
        /// <returns>A reference to a newly created BlobContainer object</returns>
        public abstract BlobContainer GetBlobContainer(string containerName);


        /// <summary>
        /// Lists the containers within the account.
        /// </summary>
        /// <returns>A list of containers</returns>
        public abstract IEnumerable<BlobContainer> ListBlobContainers();

        /// <summary>
        /// The time out for each request to the storage service.
        /// </summary>
        public TimeSpan Timeout
        {
            get;
            set;
        }

        /// <summary>
        /// The retry policy used for retrying requests
        /// </summary>
        public RetryPolicy RetryPolicy
        {
            get;
            set;
        }

        /// <summary>
        /// The base URI of the blob storage service
        /// </summary>
        public Uri BaseUri
        {
            get
            {
                return this.baseUri;
            }
        }

        /// <summary>
        /// The name of the storage account
        /// </summary>
        public string AccountName
        {
            get
            {
                return this.accountName;
            }
        }

        /// <summary>
        /// Indicates whether to use/generate path-style or host-style URIs
        /// </summary>
        public bool UsePathStyleUris
        {
            get
            {
                return this.usePathStyleUris;
            }
        }

        /// <summary>
        /// The default timeout
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes",
                          Justification = "TimeSpan is a non-mutable type")]
        public static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(30);        

        /// <summary>
        /// The default retry policy
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes",
                          Justification = "RetryPolicy is a non-mutable type")]
        public static readonly RetryPolicy DefaultRetryPolicy = RetryPolicies.NoRetry;        


        internal protected BlobStorage(Uri baseUri,
                            bool? usePathStyleUris,
                            string accountName,
                            string base64Key
                            )
        {
            this.baseUri = baseUri;
            this.accountName = accountName;
            this.Base64Key = base64Key;
            if (usePathStyleUris == null)
                this.usePathStyleUris = Utilities.StringIsIPAddress(baseUri.Host);
            else
                this.usePathStyleUris = usePathStyleUris.Value;

            Timeout = DefaultTimeout;
            RetryPolicy = DefaultRetryPolicy;
        }

       private bool usePathStyleUris;
       private Uri baseUri;
       private string accountName;
       protected internal string Base64Key
       {
           get;
           set;
       }
     }


    /// <summary>
    /// Provides definitions for some standard retry policies.
    /// </summary>
    public static class RetryPolicies
    {

        public static readonly TimeSpan StandardMinBackoff = TimeSpan.FromMilliseconds(100);
        public static readonly TimeSpan StandardMaxBackoff = TimeSpan.FromSeconds(30);
        private static readonly Random random = new Random();

        /// <summary>
        /// Policy that does no retries i.e., it just invokes <paramref name="action"/> exactly once
        /// </summary>
        /// <param name="action">The action to retry</param>
        /// <returns>The return value of <paramref name="action"/></returns>
        public static void NoRetry(Action action)
        {
            try
            {
                action();
            }
            catch (TableRetryWrapperException e)
            {
                throw e.InnerException;
            }
        }

        /// <summary>
        /// Policy that retries a specified number of times with a specified fixed time interval between retries
        /// </summary>
        /// <param name="numberOfRetries">The number of times to retry. Should be a non-negative number</param>
        /// <param name="intervalBetweenRetries">The time interval between retries. Use TimeSpan.Zero to specify immediate
        /// retries</param>
        /// <returns></returns>
        /// <remarks>When <paramref name="numberOfRetries"/> is 0 and <paramref name="intervalBetweenRetries"/> is
        /// TimeSpan.Zero this policy is equivalent to the NoRetry policy</remarks>
        public static RetryPolicy RetryN(int numberOfRetries, TimeSpan intervalBetweenRetries)
        {
            return new RetryPolicy((Action action) =>
            {
                RetryNImpl(action, numberOfRetries, intervalBetweenRetries);
            }
            );
        }

        /// <summary>
        /// Policy that retries a specified number of times with a randomized exponential backoff scheme
        /// </summary>
        /// <param name="numberOfRetries">The number of times to retry. Should be a non-negative number.</param>
        /// <param name="deltaBackoff">The multiplier in the exponential backoff scheme</param>
        /// <returns></returns>
        /// <remarks>For this retry policy, the minimum amount of milliseconds between retries is given by the 
        /// StandardMinBackoff constant, and the maximum backoff is predefined by the StandardMaxBackoff constant. 
        /// Otherwise, the backoff is calculated as random(2^currentRetry) * deltaBackoff.</remarks>
        public static RetryPolicy RetryExponentialN(int numberOfRetries, TimeSpan deltaBackoff)
        {
            return new RetryPolicy((Action action) =>
            {
                RetryExponentialNImpl(action, numberOfRetries, StandardMinBackoff, StandardMaxBackoff, deltaBackoff);
            }
            );
        }

        /// <summary>
        /// Policy that retries a specified number of times with a randomized exponential backoff scheme
        /// </summary>
        /// <param name="numberOfRetries">The number of times to retry. Should be a non-negative number</param>
        /// <param name="deltaBackoff">The multiplier in the exponential backoff scheme</param>
        /// <param name="minBackoff">The minimum backoff interval</param>
        /// <param name="maxBackoff">The maximum backoff interval</param>
        /// <returns></returns>
        /// <remarks>For this retry policy, the minimum amount of milliseconds between retries is given by the 
        /// minBackoff parameter, and the maximum backoff is predefined by the maxBackoff parameter. 
        /// Otherwise, the backoff is calculated as random(2^currentRetry) * deltaBackoff.</remarks>
        public static RetryPolicy RetryExponentialN(int numberOfRetries, TimeSpan minBackoff, TimeSpan maxBackoff, TimeSpan deltaBackoff)
        {
            if (minBackoff > maxBackoff)
            {
                throw new ArgumentException("The minimum backoff must not be larger than the maximum backoff period.");
            }
            if (minBackoff < TimeSpan.Zero)
            {
                throw new ArgumentException("The minimum backoff period must not be negative.");
            }

            return new RetryPolicy((Action action) =>
            {
                RetryExponentialNImpl(action, numberOfRetries, minBackoff, maxBackoff, deltaBackoff);
            }
            );
        }

        #region private helper methods

        private static void RetryNImpl(Action action, int numberOfRetries, TimeSpan intervalBetweenRetries)
        {
            do
            {
                try
                {
                    action();
                    break;
                }
                catch (StorageServerException)
                {
                    if (numberOfRetries == 0)
                    {
                        throw;
                    }
                    if (intervalBetweenRetries > TimeSpan.Zero)
                    {
                        Thread.Sleep(intervalBetweenRetries);
                    }
                }
                catch (TableRetryWrapperException e)
                {
                    if (numberOfRetries == 0)
                    {
                        throw e.InnerException;
                    }
                    if (intervalBetweenRetries > TimeSpan.Zero)
                    {
                        Thread.Sleep(intervalBetweenRetries);
                    }
                }
            }
            while (numberOfRetries-- > 0);
        }

        private static void RetryExponentialNImpl(Action action, int numberOfRetries, TimeSpan minBackoff, TimeSpan maxBackoff, TimeSpan deltaBackoff)
        {
            int totalNumberOfRetries = numberOfRetries;
            TimeSpan backoff;

            // sanity check
            // this is already checked when creating the retry policy in case other than the standard settings are used
            // because this library is available in source code, the standard settings can be changed and thus we 
            // check again at this point
            if (minBackoff > maxBackoff)
            {
                throw new ArgumentException("The minimum backoff must not be larger than the maximum backoff period.");
            }
            if (minBackoff < TimeSpan.Zero)
            {
                throw new ArgumentException("The minimum backoff period must not be negative.");
            }

            do
            {
                try
                {
                    action();
                    break;
                }
                catch (StorageServerException)
                {
                    if (numberOfRetries == 0)
                    {
                        throw;
                    }  
                    backoff = CalculateCurrentBackoff(minBackoff, maxBackoff, deltaBackoff, totalNumberOfRetries - numberOfRetries);
                    Debug.Assert(backoff >= minBackoff);
                    Debug.Assert(backoff <= maxBackoff);
                    if (backoff > TimeSpan.Zero) {
                        Thread.Sleep(backoff);
                    }
                }
                catch (TableRetryWrapperException e)
                {
                    if (numberOfRetries == 0)
                    {
                        throw e.InnerException;
                    }
                    backoff = CalculateCurrentBackoff(minBackoff, maxBackoff, deltaBackoff, totalNumberOfRetries - numberOfRetries);
                    Debug.Assert(backoff >= minBackoff);
                    Debug.Assert(backoff <= maxBackoff);
                    if (backoff > TimeSpan.Zero)
                    {
                        Thread.Sleep(backoff);
                    }
                }
            }
            while (numberOfRetries-- > 0);
        }

        private static TimeSpan CalculateCurrentBackoff(TimeSpan minBackoff, TimeSpan maxBackoff, TimeSpan deltaBackoff, int curRetry)
        {
            long backoff;

            if (curRetry > 30)
            {
                backoff = maxBackoff.Ticks;
            }
            else
            {
                try
                {
                    checked
                    {
                        // only randomize the multiplier here 
                        // it would be as correct to randomize the whole backoff result
                        lock (random)
                        {
                            backoff = random.Next((1 << curRetry) + 1);
                        }
                        // Console.WriteLine("backoff:" + backoff);
                        // Console.WriteLine("random range: [0, " + ((1 << curRetry) + 1) + "]");
                        backoff *= deltaBackoff.Ticks;
                        backoff += minBackoff.Ticks;
                    }
                }
                catch (OverflowException)
                {
                    backoff = maxBackoff.Ticks;
                }
                if (backoff > maxBackoff.Ticks)
                {
                    backoff = maxBackoff.Ticks;
                }
            }
            Debug.Assert(backoff >= minBackoff.Ticks);
            Debug.Assert(backoff <= maxBackoff.Ticks);
            return TimeSpan.FromTicks(backoff);
        }

        #endregion
    }

    /// <summary>
    /// Access control for containers
    /// </summary>
    public enum ContainerAccessControl
    {
        Private,
        Public
    }

    /// <summary>
    /// The blob container class.
    /// Used to access and enumerate blobs in the container.
    /// Storage key credentials are needed to access private blobs but not for public blobs.
    /// </summary>
    public abstract class BlobContainer
    {
        /// <summary>
        /// Use this constructor to access private blobs.
        /// </summary>
        /// <param name="baseUri">The base Uri for the storage endpoint</param>
        /// <param name="accountName">Name of the storage account</param>
        /// <param name="containerName">Name of the container</param>
        internal protected BlobContainer(Uri baseUri, string accountName,  string containerName)
            : this(baseUri, true, accountName, containerName, DateTime.MinValue)
        {}

        /// <summary>
        /// Use this constructor to access private blobs.
        /// </summary>
        /// <param name="baseUri">The base Uri for the storage endpoint</param>
        /// <param name="usePathStyleUris">
        /// If true, path-style URIs (http://baseuri/accountname/containername/objectname) are used and if false 
        /// host-style URIs (http://accountname.baseuri/containername/objectname) are used, where baseuri is the 
        /// URI of the service
        /// </param>
        /// <param name="accountName">Name of the storage account</param>
        /// <param name="containerName">Name of the container</param>
        /// <param name="lastModified">Date of last modification</param>
        internal protected BlobContainer(Uri baseUri, bool usePathStyleUris, string accountName, string containerName, DateTime lastModified)
        {
            if (!Utilities.IsValidContainerOrQueueName(containerName))
            {
                throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "The specified container name \"{0}\" is not valid!" +
                            "Please choose a name that conforms to the naming conventions for containers!", containerName));
            }
            this.baseUri = baseUri;
            this.usePathStyleUris = usePathStyleUris;
            this.accountName = accountName;
            this.containerName = containerName;
            this.Timeout = BlobStorage.DefaultTimeout;
            this.RetryPolicy = BlobStorage.DefaultRetryPolicy;
            this.LastModifiedTime = lastModified;
        }


        /// <summary>
        /// The time out for each request to the storage service.
        /// </summary>
        public TimeSpan Timeout
        {
            get;
            set;
        }

        /// <summary>
        /// The retry policy used for retrying requests
        /// </summary>
        public RetryPolicy RetryPolicy
        {
            get;
            set;
        }

        /// <summary>
        /// The base URI of the blob storage service
        /// </summary>
        public Uri BaseUri
        {
            get
            {
                return this.baseUri;
            }
        }

        /// <summary>
        /// The name of the storage account
        /// </summary>
        public string AccountName
        {
            get
            {
                return this.accountName;
            }
        }

        /// <summary>
        /// The name of the blob container.
        /// </summary>
        public string ContainerName
        {
            get
            {
                return this.containerName;
            }
        }

        /// <summary>
        /// Indicates whether to use/generate path-style or host-style URIs
        /// </summary>
        public bool UsePathStyleUris
        {
            get
            {
                return this.usePathStyleUris;
            }
        }

        /// <summary>
        /// The URI of the container
        /// </summary>
        public abstract Uri ContainerUri
        {
            get;
        }

        /// <summary>
        /// The timestamp for last modification of container.
        /// </summary>
        public DateTime LastModifiedTime
        {
            get;
            protected set;
        }

        /// <summary>
        /// Create the container if it does not exist.
        /// The container is created with private access control and no metadata.
        /// </summary>
        /// <returns>true if the container was created. false if the container already exists</returns>
        public abstract bool CreateContainer();

        /// <summary>
        /// Create the container with the specified metadata and access control if it does not exist
        /// </summary>
        /// <param name="metadata">The metadata for the container. Can be null to indicate no metadata</param>
        /// <param name="accessControl">The access control (public or private) with which to create the container</param>
        /// <returns>true if the container was created. false if the container already exists</returns>
        public abstract bool CreateContainer(NameValueCollection metadata, ContainerAccessControl accessControl); 

        /// <summary>
        /// Check if the blob container exists
        /// </summary>
        /// <returns>true if the container exists, false otherwise.</returns>
        public abstract bool DoesContainerExist();

        /// <summary>
        /// Get the properties for the container if it exists.
        /// </summary>
        /// <returns>The properties for the container if it exists, null otherwise</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate",
            Justification="The method makes a call to the blob service")]
        public abstract ContainerProperties GetContainerProperties();

        /// <summary>
        /// Get the access control permissions associated with the container.
        /// </summary>
        /// <returns></returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate",
            Justification = "The method makes a call to the blob service")]
        public abstract ContainerAccessControl GetContainerAccessControl();

        /// <summary>
        /// Set the access control permissions associated with the container.
        /// </summary>
        /// <param name="acl">The permission to set</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate",
            Justification = "The method makes a call to the blob service")]
        public abstract void SetContainerAccessControl(ContainerAccessControl acl);

        /// <summary>
        /// Deletes the current container.
        /// </summary>
        public abstract bool DeleteContainer();

        /// <summary>
        /// Check if the blob container exists
        /// </summary>
        /// <param name="blobName">Name of the BLOB.</param>
        /// <returns>true if the blob exists, false otherwise.</returns>
        public abstract bool DoesBlobExist(string blobName);

        /// <summary>
        /// Create a new blob or overwrite an existing blob.
        /// </summary>
        /// <param name="blobProperties">The properties of the blob</param>
        /// <param name="blobContents">The contents of the blob</param>
        /// <param name="overwrite">Should this request overwrite an existing blob ?</param>
        /// <returns>true if the blob was created. false if the blob already exists and <paramref name="overwrite"/>was set to false</returns>
        /// <remarks>The LastModifiedTime property of <paramref name="blobProperties"/> is set as a result of this call.
        /// This method also has an effect on the ETag values that are managed by the service.</remarks>
        public abstract bool CreateBlob(BlobProperties blobProperties, BlobContents blobContents, bool overwrite);

        /// <summary>
        /// Updates an existing blob if it has not been modified since the specified time which is typically
        /// the last modified time of the blob when you retrieved it.
        /// Use this method to implement optimistic concurrency by avoiding clobbering changes to the blob
        /// made by another writer.
        /// </summary>
        /// <param name="blob">The properties of the blob. This object should be one previously
        /// obtained from a call to GetBlob or GetBlobProperties and have its LastModifiedTime property set.</param>
        /// <param name="contents">The contents of the blob. The contents of the blob should be readable</param>
        /// <returns>true if the blob was updated. false if the blob has changed since the last time</returns>
        /// <remarks>The LastModifiedTime property of <paramref name="blob"/> is set as a result of this call.
        /// This method also has an effect on the ETag values that are managed by the service if the update was 
        /// successful.</remarks>
        public abstract bool UpdateBlobIfNotModified(BlobProperties blob, BlobContents contents);

        /// <summary>
        /// Get the blob contents and properties if the blob exists
        /// </summary>
        /// <param name="name">The name of the blob</param>
        /// <param name="blobContents">Object in which the contents are returned.
        /// This object should contain a writable stream or should be a default constructed object.</param>
        /// <param name="transferAsChunks">Should the blob be gotten in pieces. This requires more round-trips, but will retry smaller pieces in case of failure.</param>
        /// <returns>The properties of the blob if the blob exists.</returns>
        public abstract BlobProperties GetBlob(string name, BlobContents blobContents, bool transferAsChunks);

        /// <summary>
        /// Gets the blob contents and properties if the blob has not been modified since the time specified.
        /// Use this method if you have cached the contents of a blob and want to avoid retrieving the blob
        /// if it has not changed since the last time you retrieved it.
        /// </summary>
        /// <param name="blobProperties">The properties of the blob obtained from an earlier call to GetBlob. This
        /// parameter is updated by the call if the blob has been modified</param>
        /// <param name="blobContents">Contains the stream to which the contents of the blob are written if it has been
        /// modified</param>
        /// <param name="transferAsChunks">Should the blob be gotten in pieces. This requires more round-trips, but will retry smaller pieces in case of failure.</param>
        /// <returns>true if the blob has been modified, false otherwise</returns>
        public abstract bool GetBlobIfModified(BlobProperties blobProperties, BlobContents blobContents, bool transferAsChunks);

        /// <summary>
        /// Get the properties of the blob if it exists.
        /// This method is also the simplest way to check if a blob exists.
        /// </summary>
        /// <param name="name">The name of the blob</param>
        /// <returns>The properties of the blob if it exists. null otherwise.
        /// The properties for the contents of the blob are not set</returns>
        public abstract BlobProperties GetBlobProperties(string name);

        /// <summary>
        /// Set the metadata of an existing blob.
        /// </summary>
        /// <param name="blobProperties">The blob properties object whose metadata is to be updated</param>
        public abstract void UpdateBlobMetadata(BlobProperties blobProperties);

        /// <summary>
        /// Set the metadata of an existing blob if it has not been modified since it was last retrieved.
        /// </summary>
        /// <param name="blobProperties">The blob properties object whose metadata is to be updated.
        /// Typically obtained by a previous call to GetBlob or GetBlobProperties</param>
        /// <returns>true if the blob metadata was updated. false if it was not updated because the blob
        /// has been modified</returns>
        public abstract bool UpdateBlobMetadataIfNotModified(BlobProperties blobProperties);
        
        /// <summary>
        /// Delete a blob with the given name
        /// </summary>
        /// <param name="name">The name of the blob</param>
        /// <returns>true if the blob exists and was successfully deleted, false if the blob does not exist</returns>
        public abstract bool DeleteBlob(string name);

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
        public abstract bool DeleteBlobIfNotModified(BlobProperties blob, out bool modified);

        /// <summary>
        /// Enumerates all blobs with a given prefix.
        /// </summary>
        /// <param name="prefix"></param>
        /// <param name="combineCommonPrefixes">If true common prefixes with "/" as seperator</param>
        /// <returns>The list of blob properties and common prefixes</returns>
        public abstract IEnumerable<object> ListBlobs(string prefix, bool combineCommonPrefixes);

        private Uri    baseUri;
        private string accountName;
        private string containerName;
        private bool usePathStyleUris;
    }

    /// <summary>
    /// The properties of a blob.
    /// No member of this class makes a storage service request.
    /// </summary>
    public class BlobProperties
    {
        /// <summary>
        /// Construct a new BlobProperties object
        /// </summary>
        /// <param name="name">The name of the blob</param>
        public BlobProperties(string name)
        {
            Name = name;
        }


        /// <summary>
        /// Name of the blob
        /// </summary>
        public string Name { get; internal set; }

        /// <summary>
        /// URI of the blob
        /// </summary>
        public Uri Uri { get; internal set; }

        /// <summary>
        /// Content encoding of the blob if it set, null otherwise.
        /// </summary>
        public string ContentEncoding { get; set; }

        /// <summary>
        /// Content Type of the blob if it is set, null otherwise.
        /// </summary>
        public string ContentType { get; set; }

        /// <summary>
        /// Content Language of the blob if it is set, null otherwise.
        /// </summary>
        public string ContentLanguage { get; set; }

        /// <summary>
        /// The length of the blob content, null otherwise.
        /// </summary>
        public long ContentLength { get; internal set; }

        /// <summary>
        /// Metadata for the blob in the form of name-value pairs.
        /// </summary>
        public NameValueCollection Metadata { get; set;}

        /// <summary>
        /// The last modified time for the blob. 
        /// </summary>
        public DateTime LastModifiedTime { get; internal set; }

        /// <summary>
        /// The ETag of the blob. This is an identifier assigned to the blob by the storage service
        /// and is used to distinguish contents of two blobs (or versions of the same blob).
        /// </summary>
        public string ETag { get; internal set; }

        internal void Assign(BlobProperties other)
        {
            Name = other.Name;
            Uri = other.Uri;
            ContentEncoding = other.ContentEncoding;
            ContentLanguage = other.ContentLanguage;
            ContentLength = other.ContentLength;
            ContentType = other.ContentType;
            ETag = other.ETag;
            LastModifiedTime = other.LastModifiedTime;
            Metadata = (other.Metadata != null ? new NameValueCollection(other.Metadata) : null) ;
        }
    }

    /// <summary>
    /// The properties of a container.
    /// No member of this class makes a storage service request.
    /// </summary>
    public class ContainerProperties
    {
        public ContainerProperties(string name)
        {
            Name = name;
        }

        public string Name
        {
            get;
            internal set;
        }
        public DateTime LastModifiedTime
        {
            get;
            internal set;
        }

        public string ETag
        {
            get;
            internal set;
        }

        public Uri Uri
        {
            get;
            internal set;
        }

        public NameValueCollection Metadata
        {
            get;
            internal set;
        }
    }


    /// <summary>
    /// The contents of the Blob in various forms.
    /// </summary>
    public class BlobContents
    {
        /// <summary>
        /// Construct a new BlobContents object from a stream.
        /// </summary>
        /// <param name="stream">The stream to/from which blob contents are written/read. The
        /// stream should be seekable in order for requests to be retried.</param>
        public BlobContents(Stream stream)
        {
            this.stream = stream;
        }

        /// <summary>
        /// Construct a new BlobContents object from a byte array.
        /// </summary>
        /// <param name="value">The byte array to/from which contents are written/read.</param>
        public BlobContents(byte[] value)
        {
            this.bytes = value;
            this.stream = new MemoryStream(value, false);
        }

        /// <summary>
        /// Get the contents of a blob as a stream.
        /// </summary>
        public Stream AsStream
        {
            get
            {
                return stream;
            }

        }

        /// <summary>
        /// Get the contents of a blob as a byte array.
        /// </summary>
        public byte[] AsBytes()
        {
            if (bytes != null)
                return bytes;
            if (stream != null)
            {
                stream.Seek(0, SeekOrigin.Begin);
                bytes = new byte[stream.Length];
                int n = 0;
                int offset = 0;
                do
                {
                    n = stream.Read(bytes, offset, bytes.Length - offset);
                    offset += n;

                } while (n > 0);
            }
            return bytes;
        }

        private Stream stream;
        private byte[] bytes;
    }
    #endregion
}