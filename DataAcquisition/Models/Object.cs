﻿using System;
using System.Collections.Generic;

namespace DataAcquisition.Models;

public partial class Object
{
    public Guid Id { get; set; }

    public string? BucketId { get; set; }

    public string? Name { get; set; }

    public Guid? Owner { get; set; }

    public DateTime? CreatedAt { get; set; }

    public DateTime? UpdatedAt { get; set; }

    public DateTime? LastAccessedAt { get; set; }

    public string? Metadata { get; set; }

    public string[]? PathTokens { get; set; }

    public virtual Bucket? Bucket { get; set; }

    public virtual User1? OwnerNavigation { get; set; }
}
