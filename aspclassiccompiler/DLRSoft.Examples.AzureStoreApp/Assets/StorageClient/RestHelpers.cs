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
// <copyright file="RestHelpers.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net;
using System.Diagnostics;
using System.Globalization;
using System.Collections.Specialized;
using System.Web;
using System.Xml;
using System.Text.RegularExpressions;

namespace Microsoft.Samples.ServiceHosting.StorageClient
{
    namespace StorageHttpConstants
    {

        internal static class ConstChars
        {
            internal const string Linefeed = "\n";
            internal const string CarriageReturnLinefeed = "\r\n";
            internal const string Colon = ":";
            internal const string Comma = ",";
            internal const string Slash = "/";
            internal const string BackwardSlash = @"\";
            internal const string Space = " ";
            internal const string Ampersand = "&";
            internal const string QuestionMark = "?";
            internal const string Equal = "=";
            internal const string Bang = "!";
            internal const string Star = "*";
            internal const string Dot = ".";
        }

        internal static class RequestParams
        {
            internal const string NumOfMessages = "numofmessages";
            internal const string VisibilityTimeout = "visibilitytimeout";
            internal const string PeekOnly = "peekonly";
            internal const string MessageTtl = "messagettl";
            internal const string Messages = "messages";
            internal const string PopReceipt = "popreceipt";
        }

        internal static class QueryParams
        {
            internal const string SeparatorForParameterAndValue = "=";
            internal const string QueryParamTimeout = "timeout";
            internal const string QueryParamComp = "comp";

            // Other query string parameter names
            internal const string QueryParamBlockId = "blockid";
            internal const string QueryParamPrefix = "prefix";
            internal const string QueryParamMarker = "marker";
            internal const string QueryParamMaxResults = "maxresults";
            internal const string QueryParamDelimiter = "delimiter";
            internal const string QueryParamModifiedSince = "modifiedsince";
        }

        internal static class CompConstants
        {
            internal const string Metadata = "metadata";
            internal const string List = "list";
            internal const string BlobList = "bloblist";
            internal const string BlockList = "blocklist";
            internal const string Block = "block";
            internal const string Acl = "acl";
        }

        internal static class XmlElementNames
        {
            internal const string BlockList = "BlockList";
            internal const string Block = "Block";
            internal const string EnumerationResults = "EnumerationResults";
            internal const string Prefix = "Prefix";
            internal const string Marker = "Marker";
            internal const string MaxResults = "MaxResults";
            internal const string Delimiter = "Delimiter";
            internal const string NextMarker = "NextMarker";
            internal const string Containers = "Containers";
            internal const string Container = "Container";
            internal const string ContainerName = "Name";
            internal const string ContainerNameAttribute = "ContainerName";
            internal const string AccountNameAttribute = "AccountName";
            internal const string LastModified = "LastModified";
            internal const string Etag = "Etag";
            internal const string Url = "Url";
            internal const string CommonPrefixes = "CommonPrefixes";
            internal const string ContentType = "ContentType";
            internal const string ContentEncoding = "ContentEncoding";
            internal const string ContentLanguage = "ContentLanguage";
            internal const string Size = "Size";
            internal const string Blobs = "Blobs";
            internal const string Blob = "Blob";
            internal const string BlobName = "Name";
            internal const string BlobPrefix = "BlobPrefix";
            internal const string BlobPrefixName = "Name";
            internal const string Name = "Name";
            internal const string Queues = "Queues";
            internal const string Queue = "Queue";
            internal const string QueueName = "QueueName";
            internal const string QueueMessagesList = "QueueMessagesList";
            internal const string QueueMessage = "QueueMessage";
            internal const string MessageId = "MessageId";
            internal const string PopReceipt = "PopReceipt";
            internal const string InsertionTime = "InsertionTime";
            internal const string ExpirationTime = "ExpirationTime";
            internal const string TimeNextVisible = "TimeNextVisible";
            internal const string MessageText = "MessageText";

            // Error specific constants
            internal const string ErrorRootElement = "Error";
            internal const string ErrorCode = "Code";
            internal const string ErrorMessage = "Message";
            internal const string ErrorException = "ExceptionDetails";
            internal const string ErrorExceptionMessage = "ExceptionMessage";
            internal const string ErrorExceptionStackTrace = "StackTrace";
            internal const string AuthenticationErrorDetail = "AuthenticationErrorDetail";

            //The following are for table error messages
            internal const string DataWebMetadataNamespace = "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata";
            internal const string TableErrorCodeElement = "code";
            internal const string TableErrorMessageElement = "message";
        }

        internal static class HeaderNames
        {
            internal const string PrefixForStorageProperties = "x-ms-prop-";
            internal const string PrefixForMetadata = "x-ms-meta-";
            internal const string PrefixForStorageHeader = "x-ms-";
            internal const string PrefixForTableContinuation = "x-ms-continuation-";

            //
            // Standard headers...
            //
            internal const string ContentLanguage = "Content-Language";
            internal const string ContentLength = "Content-Length";
            internal const string ContentType = "Content-Type";
            internal const string ContentEncoding = "Content-Encoding";
            internal const string ContentMD5 = "Content-MD5";
            internal const string ContentRange = "Content-Range";
            internal const string LastModifiedTime = "Last-Modified";
            internal const string Server = "Server";
            internal const string Allow = "Allow";
            internal const string ETag = "ETag";
            internal const string Range = "Range";
            internal const string Date = "Date";
            internal const string Authorization = "Authorization";
            internal const string IfModifiedSince = "If-Modified-Since";
            internal const string IfUnmodifiedSince = "If-Unmodified-Since";
            internal const string IfMatch = "If-Match";
            internal const string IfNoneMatch = "If-None-Match";
            internal const string IfRange = "If-Range";
            internal const string NextPartitionKey = "NextPartitionKey";
            internal const string NextRowKey = "NextRowKey";
            internal const string NextTableName = "NextTableName";

            //
            // Storage specific custom headers...
            //
            internal const string StorageDateTime = PrefixForStorageHeader + "date";
            internal const string PublicAccess = PrefixForStorageProperties + "publicaccess";
            internal const string StorageRange = PrefixForStorageHeader + "range";

            internal const string CreationTime = PrefixForStorageProperties + "creation-time";
            internal const string ForceUpdate = PrefixForStorageHeader + "force-update";            
            internal const string ApproximateMessagesCount = PrefixForStorageHeader + "approximate-messages-count";
            internal const string Version = PrefixForStorageHeader + "version";     
        }

        internal static class HeaderValues
        {
            internal const string ContentTypeXml = "application/xml";

            /// <summary>
            /// This is the default content-type xStore uses when no content type is specified
            /// </summary>
            internal const string DefaultContentType = "application/octet-stream";

            // The Range header value is "bytes=start-end", both start and end can be empty
            internal const string RangeHeaderFormat = "bytes={0}-{1}";

        }

        internal static class AuthenticationSchemeNames
        {
            internal const string SharedKeyAuthSchemeName = "SharedKey";
            internal const string SharedKeyLiteAuthSchemeName = "SharedKeyLite";
        }

        internal static class HttpMethod
        {
            internal const string Get = "GET";
            internal const string Put = "PUT";
            internal const string Post = "POST";
            internal const string Head = "HEAD";
            internal const string Delete = "DELETE";
            internal const string Trace = "TRACE";
            internal const string Options = "OPTIONS";
            internal const string Connect = "CONNECT";
        }

        internal static class BlobBlockConstants
        {
            internal const int KB = 1024;
            internal const int MB = 1024 * KB;
            /// <summary>
            /// When transmitting a blob that is larger than this constant, this library automatically
            /// transmits the blob as individual blocks. I.e., the blob is (1) partitioned
            /// into separate parts (these parts are called blocks) and then (2) each of the blocks is 
            /// transmitted separately.
            /// The maximum size of this constant as supported by the real blob storage service is currently 
            /// 64 MB; the development storage tool currently restricts this value to 2 MB.
            /// Setting this constant can have a significant impact on the performance for uploading or
            /// downloading blobs.
            /// As a general guideline: If you run in a reliable environment increase this constant to reduce
            /// the amount of roundtrips. In an unreliable environment keep this constant low to reduce the 
            /// amount of data that needs to be retransmitted in case of connection failures.
            /// </summary>
            internal const long MaximumBlobSizeBeforeTransmittingAsBlocks = 2 * MB;
            /// <summary>
            /// The size of a single block when transmitting a blob that is larger than the 
            /// MaximumBlobSizeBeforeTransmittingAsBlocks constant (see above).
            /// The maximum size of this constant is currently 4 MB; the development storage 
            /// tool currently restricts this value to 1 MB.
            /// Setting this constant can have a significant impact on the performance for uploading or 
            /// downloading blobs.
            /// As a general guideline: If you run in a reliable environment increase this constant to reduce
            /// the amount of roundtrips. In an unreliable environment keep this constant low to reduce the 
            /// amount of data that needs to be retransmitted in case of connection failures.
            /// </summary>
            internal const long BlockSize = 1 * MB;            
        }

        internal static class ListingConstants
        {
            internal const int MaxContainerListResults = 100;
            internal const int MaxBlobListResults = 100;
            internal const int MaxQueueListResults = 50;
            internal const int MaxTableListResults = 50;
        }

        /// <summary>
        /// Contains regular expressions for checking whether container and table names conform
        /// to the rules of the storage REST protocols.
        /// </summary>
        public static class RegularExpressionStrings
        {
            /// <summary>
            /// Container or queue names that match against this regular expression are valid.
            /// </summary>
            public const string ValidContainerNameRegex = @"^([a-z]|\d){1}([a-z]|-|\d){1,61}([a-z]|\d){1}$";

            /// <summary>
            /// Table names that match against this regular expression are valid.
            /// </summary>
            public const string ValidTableNameRegex = @"^([a-z]|[A-Z]){1}([a-z]|[A-Z]|\d){2,62}$";
        }

        internal static class StandardPortalEndpoints
        {
            internal const string BlobStorage = "blob";
            internal const string QueueStorage = "queue";
            internal const string TableStorage = "table";
            internal const string StorageHostSuffix = ".core.windows.net";
            internal const string BlobStorageEndpoint = BlobStorage + StorageHostSuffix;
            internal const string QueueStorageEndpoint = QueueStorage + StorageHostSuffix;
            internal const string TableStorageEndpoint = TableStorage + StorageHostSuffix;
        }
    }

    internal static partial class Utilities
    {
        internal static HttpWebRequest CreateHttpRequest(Uri uri, string httpMethod, TimeSpan timeout)
        {
            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(uri);
            request.Timeout = (int)timeout.TotalMilliseconds;
            request.ReadWriteTimeout = (int)timeout.TotalMilliseconds;
            request.Method = httpMethod;
            request.ContentLength = 0;
            request.Headers.Add(StorageHttpConstants.HeaderNames.StorageDateTime,
                                Utilities.ConvertDateTimeToHttpString(DateTime.UtcNow));
            return request;
        }

        /// <summary>
        /// Converts the date time to a valid string form as per HTTP standards
        /// </summary>
        internal static string ConvertDateTimeToHttpString(DateTime dateTime)
        {
            // On the wire everything should be represented in UTC. This assert will catch invalid callers who
            // are violating this rule.
            Debug.Assert(dateTime == DateTime.MaxValue || dateTime == DateTime.MinValue || dateTime.Kind == DateTimeKind.Utc);

            // 'R' means rfc1123 date which is what our server uses for all dates...
            // It will be in the following format:
            // Sun, 28 Jan 2008 12:11:37 GMT
            return dateTime.ToString("R", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Parse a string having the date time information in acceptable formats according to HTTP standards
        /// </summary>
        internal static bool TryGetDateTimeFromHttpString(string dateString, out DateTime? result)
        {
            DateTime dateTime;
            result = null;

            // 'R' means rfc1123 date which is the preferred format used in HTTP
            bool parsed = DateTime.TryParseExact(dateString, "R", null, DateTimeStyles.None, out dateTime);
            if (parsed)
            {
                // For some reason, format string "R" makes the DateTime.Kind as Unspecified while it's actually
                // Utc. Specifying DateTimeStyles.AssumeUniversal also doesn't make the difference. If we also
                // specify AdjustToUniversal it works as expected but we don't really want Parse to adjust 
                // things automatically.
                result = DateTime.SpecifyKind(dateTime, DateTimeKind.Utc);
                return true;
            }

            return false;
        }


        /// <summary>
        /// Copies from one stream to another
        /// </summary>
        /// <param name="sourceStream">The stream to copy from</param>
        /// <param name="destinationStream">The stream to copy to</param>
        internal static long CopyStream(Stream sourceStream, Stream destinationStream)
        {
            const int BufferSize = 0x10000;
            byte[] buffer = new byte[BufferSize];
            int n = 0;
            long totalRead = 0;
            do
            {
                n = sourceStream.Read(buffer, 0, BufferSize);
                if (n > 0)
                {
                    totalRead += n;
                    destinationStream.Write(buffer, 0, n);
                }
                
            } while (n > 0);
            return totalRead;
        }

        internal static void CopyStream(Stream sourceStream, Stream destinationStream, long length)
        {
            const int BufferSize = 0x10000;
            byte[] buffer = new byte[BufferSize];
            int n = 0;
            long amountLeft = length;           

            do
            {
                amountLeft -= n;
                n = sourceStream.Read(buffer, 0, (int)Math.Min(BufferSize, amountLeft));
                if (n > 0)
                {
                    destinationStream.Write(buffer, 0, n);
                }

            } while (n > 0);
        }

        internal static int CopyStreamToBuffer(Stream sourceStream, byte[] buffer, int bytesToRead)
        {
            int n = 0;
            int amountLeft = bytesToRead;
            do
            {
                n = sourceStream.Read(buffer, bytesToRead - amountLeft, amountLeft);
                amountLeft -= n;
            } while (n > 0);
            return bytesToRead - amountLeft;
        }

        internal static Uri CreateRequestUri(
                                Uri baseUri,
                                bool usePathStyleUris,
                                string accountName,
                                string containerName,
                                string blobName,
                                TimeSpan Timeout,
                                NameValueCollection queryParameters,
                                out ResourceUriComponents uriComponents
                                )
        {
            uriComponents = 
                new ResourceUriComponents(accountName, containerName, blobName);
            Uri uri = HttpRequestAccessor.ConstructResourceUri(baseUri, uriComponents, usePathStyleUris);

            if (queryParameters != null)
            {
                UriBuilder builder = new UriBuilder(uri);

                if (queryParameters.Get(StorageHttpConstants.QueryParams.QueryParamTimeout) == null)
                {
                    queryParameters.Add(StorageHttpConstants.QueryParams.QueryParamTimeout,
                    Timeout.TotalSeconds.ToString(CultureInfo.InvariantCulture));
                }

                StringBuilder sb = new StringBuilder();
                bool firstParam = true;
                foreach (string queryKey in queryParameters.AllKeys)
                {
                    if (!firstParam)
                        sb.Append("&");
                    sb.Append(HttpUtility.UrlEncode(queryKey));
                    sb.Append('=');
                    sb.Append(HttpUtility.UrlEncode(queryParameters[queryKey]));
                    firstParam = false;
                }

                if (sb.Length > 0)
                {
                    builder.Query = sb.ToString();
                }
                return builder.Uri;
            }
            else
            {
                return uri;
            }
        }

        internal static bool StringIsIPAddress(string address)
        {
            IPAddress outIPAddress;

            return IPAddress.TryParse(address, out outIPAddress);
        }

        internal static void AddMetadataHeaders(HttpWebRequest request, NameValueCollection metadata)
        {
            foreach (string key in metadata.Keys)
            {
                request.Headers.Add(
                    StorageHttpConstants.HeaderNames.PrefixForMetadata + key,
                    metadata[key]
                    );
            }
        }

        internal static bool IsValidTableName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }
            Regex reg = new Regex(StorageHttpConstants.RegularExpressionStrings.ValidTableNameRegex);
            if (reg.IsMatch(name))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        internal static bool IsValidContainerOrQueueName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }
            Regex reg = new Regex(StorageHttpConstants.RegularExpressionStrings.ValidContainerNameRegex);
            if (reg.IsMatch(name))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}