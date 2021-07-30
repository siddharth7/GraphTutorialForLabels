# GraphTutorialForLabels

Code snippet to fetch labels in a drive 

```
public class SensitivityLabel
{
    public string displayName
    {
        get;
        set;
    }
    public string id
    {
        get;
        set;
    }
    public bool protectionEnabled
    {
        get;
        set;
    }
}
    
public static async Task GetAllFilesAsync()
{
    IDriveItemSearchCollectionPage driveItems = await graphClient.Me.Drive.Root.Search("").Request().GetAsync();
    List<string> driveItemIds = new List<string>();

    var driveItemsIterator = PageIterator<DriveItem>
    .CreatePageIterator(
    graphClient,
    driveItems,
    (item) =>
        {
            if (item.File != null)
            {
                driveItemIds.Add(item.Id);
            }
            return true;
        },
    (req) =>
        {
            return req;
        }
    );

    await driveItemsIterator.IterateAsync();

    foreach (var itemId in driveItemIds)
    {
        var driveItemInfo = await graphClient.Me.Drive.Items[itemId].Request().Select("Name,WebUrl,sensitivitylabel").GetAsync();
        object sensitivityLabelData;
        driveItemInfo.AdditionalData.TryGetValue("sensitivityLabel", out sensitivityLabelData);
        var sensitivityLabel = JsonSerializer.Deserialize<SensitivityLabel>(sensitivityLabelData.ToString());
        if(!string.IsNullOrEmpty(sensitivityLabel.id))
        {
            Console.WriteLine("File Name: {0}, File Url: {1}, Sensitivitylabel Name: {2}, SensitivitylabelId: {3}, IsProtectectionEnabled: {4}",
                driveItemInfo.Name, driveItemInfo.WebUrl, sensitivityLabel.displayName, sensitivityLabel.id, sensitivityLabel.protectionEnabled);
        }
    }
}
```
