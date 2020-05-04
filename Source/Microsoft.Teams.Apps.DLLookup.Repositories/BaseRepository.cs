// <copyright file="BaseRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Repositories
{
    using System.Collections.Generic;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Extensions.Configuration;

    /// <summary>
    /// Base repository for the data stored in the Azure Table Storage.
    /// </summary>
    /// <typeparam name="T">Entity class type.</typeparam>
    public class BaseRepository<T>
        where T : TableEntity, new()
    {
#pragma warning disable CS0649
        private readonly string defaultPartitionKey;
#pragma warning disable CS0649

        /// <summary>
        /// Initializes a new instance of the <see cref="BaseRepository{T}"/> class.
        /// </summary>
        /// <param name="configuration">Singleton instance of application configuration.</param>
        /// <param name="tableName">The name of the table in Azure Table Storage.</param>
        public BaseRepository(
            IConfiguration configuration,
            string tableName)
        {
            string storageAccountConnectionString = configuration["StorageAccountConnectionString"];
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(storageAccountConnectionString);
            CloudTableClient tableClient = storageAccount.CreateCloudTableClient();
            this.Table = tableClient.GetTableReference(tableName);
            if (!this.Table.Exists())
            {
                this.Table.CreateIfNotExists();
            }
        }

        /// <summary>
        /// Gets cloud table instance.
        /// </summary>
        public CloudTable Table { get; }

        /// <summary>
        /// Create or update an entity in the table storage.
        /// </summary>
        /// <param name="entity">Entity to be created or updated.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task CreateOrUpdateAsync(T entity)
        {
            TableOperation operation = TableOperation.InsertOrReplace(entity);

            await this.Table.ExecuteAsync(operation);
        }

        /// <summary>
        /// Delete an entity in the table storage.
        /// </summary>
        /// <param name="entity">Entity to be deleted.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task DeleteAsync(T entity)
        {
            TableOperation operation = TableOperation.Delete(entity);

            await this.Table.ExecuteAsync(operation);
        }

        /// <summary>
        /// Get an entity by the keys in the table storage.
        /// </summary>
        /// <param name="partitionKey">The partition key of the entity.</param>
        /// <param name="rowKey">The row key to the entity.</param>
        /// <returns>The entity matching the keys.</returns>
        public async Task<T> GetAsync(string partitionKey, string rowKey)
        {
            TableOperation operation = TableOperation.Retrieve<T>(partitionKey.ToLower(), rowKey.ToLower());

            TableResult result = await this.Table.ExecuteAsync(operation);

            return result.Result as T;
        }

        /// <summary>
        /// Get all data entities from the table storage in a partition.
        /// </summary>
        /// <param name="partition">Partition key value.</param>
        /// <param name="count">The max number of desired entities.</param>
        /// <returns>All data entities.</returns>
        public async Task<IEnumerable<T>> GetAllAsync(string partition = null, int? count = null)
        {
            string partitionKeyFilter = this.GetPartitionKeyFilter(partition);

            TableQuery<T> query = new TableQuery<T>().Where(partitionKeyFilter);

            IList<T> entities = await this.ExecuteQueryAsync(query, count);

            return entities;
        }

        /// <summary>
        /// Forms filter condition with Partition key.
        /// </summary>
        /// <param name="partition">Partition key value.</param>
        /// <returns>Filter condition with Partition key.</returns>
        private string GetPartitionKeyFilter(string partition)
        {
            string filter = TableQuery.GenerateFilterCondition(
                nameof(TableEntity.PartitionKey),
                QueryComparisons.Equal,
                string.IsNullOrWhiteSpace(partition) ? this.defaultPartitionKey : partition.ToLower());
            return filter;
        }

        /// <summary>
        /// Execute Table query.
        /// </summary>
        /// <param name="query">query to filter records.</param>
        /// <param name="count">Optional parameter. Maximum number of desired entities.</param>
        /// <param name="ct">Cancellation token details.</param>
        /// <returns>Result of the asynchronous operation.</returns>
        private async Task<IList<T>> ExecuteQueryAsync(
            TableQuery<T> query,
            int? count = null,
            CancellationToken ct = default)
        {
            query.TakeCount = count;

            try
            {
                List<T> result = new List<T>();
                TableContinuationToken token = null;

                do
                {
                    TableQuerySegment<T> seg = await this.Table.ExecuteQuerySegmentedAsync<T>(query, token);
                    token = seg.ContinuationToken;
                    result.AddRange(seg);
                }
                while (token != null
                    && !ct.IsCancellationRequested
                    && (count == null || result.Count < count.Value));

                return result;
            }
            catch
            {
                throw;
            }
        }
    }
}
