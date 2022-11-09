﻿using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TabWithAdpativeCardFlow.Interfaces
{
    public interface IGraphHelper
    {
        GraphServiceClient GetDelegatedServiceClient(string _token);
    }
}
