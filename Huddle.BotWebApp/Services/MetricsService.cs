/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using Huddle.BotWebApp.Models;
using System.Linq;
using System.Threading.Tasks;

namespace Huddle.BotWebApp.Services
{
    public class MetricsService : GraphService
    {
        private string baseSPSiteUrl;

        public MetricsService(string token, string baseSPSiteUrl) :
            base(token)
        {
            this.baseSPSiteUrl = baseSPSiteUrl;
        }

        public async Task<Metric[]> GetActiveMetricsAsync(string teamId)
        {
            var site = this._graphServiceClient.Sites.Root.SiteWithPath(baseSPSiteUrl);

            var issues = await site.Lists["Issues"].Items.Request()
                .Select("Id")
                .Filter($"fields/HuddleTeamId eq '{teamId}'")
                .GetAllAsync();

            var fitler = string.Join(" or ", issues.Select(i => "fields/HuddleIssueLookupId eq " + i.Id));

            var metrics = await site.Lists["Metrics"].Items.Request()
                .Expand("fields($select=Title)")
                .Filter(fitler)
                .GetAllAsync();

            return metrics.Select(i => new Metric
            {
                Id = int.Parse(i.Id),
                Name = i.Fields.AdditionalData["Title"] as string
            }
            ).ToArray();
        }
    }
}

