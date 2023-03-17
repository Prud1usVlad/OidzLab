﻿using System;
using System.Collections.Generic;

namespace DataAcquisition.Models;

public partial class User
{
    public Guid Id { get; set; }

    public string? Gender { get; set; }

    public int? Age { get; set; }

    public string? Country { get; set; }

    public virtual ICollection<Event> Events { get; } = new List<Event>();
}