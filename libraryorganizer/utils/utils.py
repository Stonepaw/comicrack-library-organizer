def UIThread(func):    
    def wrapper(self, *args): 
        if self.InvokeRequired:
          pass
        else:
            pass
    if len(args) == 0:    
      actiontype = Action1[object]    
    else:    
      actiontype = Action[tuple(object for x in range(len(args)+1))]    
 
    action = actiontype(fun)    
    self.dispatcher.Invoke(action, self, *args)    
      
    return wrapper

