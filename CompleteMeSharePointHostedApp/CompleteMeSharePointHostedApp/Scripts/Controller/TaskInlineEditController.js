myApp.directive('inlineEdit', ['$timeout', '$compile', '$http', '$templateCache', 'SharePointJSOMService', function ($timeout, $compile, $http, $templateCache, SharePointJSOMService) {
    return {
        scope: {
            todo: '=inlineEdit',
            handleSave: '&onSave',
            handleCancel: '&onCancel'
        },
        //NOTE: can't use this as path changes --  , templateUrl:  '/sites/test2/Complete%20Me/Pages/template.html' 
        compile: function compile(tElement, tAttrs) {
            JSRequest.EnsureSetup();
            appweburl = decodeURIComponent(JSRequest.QueryString["SPAppWebUrl"]);
            var tplURL = appweburl + '/Pages/template.html';
            templateLoader = $http.get(tplURL, { cache: $templateCache })
                .success(function (html) {
                    tElement.html(html);
                });
            return {
                //have to return link: here because when using compile: ignores link: in directive
                pre: function preLink(scope, iElement, iAttrs, controller) {
                    //this is required to load the template HTML as using compile
                    templateLoader.then(function (templateText) {
                        iElement.html($compile(tElement.html())(scope));
                    });

                    var previousValue;

                    scope.edit = function () {
                        scope.editMode = true;
                        previousValue = scope.todo;
                        initializePeoplePicker('peoplePickerDiv' + scope.todo.id, scope.todo.assignedTo, scope.todo.assignedToName);
                        $timeout(function () {
                            elm.find('input')[1].focus();
                        }, 0, false); //TODO: not working!
                    };
                    scope.save = function () {
                        $.when(SharePointJSOMService.saveTask(scope, TaskListName))
                        .done(function (jsonObject) {
                        })
                        .fail(function (err) {
                            console.info(JSON.stringify(err));
                        });
                        scope.editMode = false;
                        scope.handleSave({ value: scope.todo });
                    };
                    scope.delete = function () {
                        $.when(SharePointJSOMService.deleteTask(scope, TaskListName))
                        .done(function (jsonObject) {
                        })
                        .fail(function (err) {
                            console.info(JSON.stringify(err));
                        });
                        scope.isDeleted = true;
                        scope.handleSave({ value: scope.todo });
                    };
                    scope.cancel = function () {
                        scope.editMode = false;
                        scope.todo = previousValue;
                        scope.handleCancel({ value: scope.todo });
                    };
                    scope.bump = function (date) {
                        scope.editMode = false;
                        scope.todo.dueDate = date;

                        $.when(SharePointJSOMService.bumpTask(scope, TaskListName))
                        .done(function (jsonObject) {
                        })
                        .fail(function (err) {
                            console.info(JSON.stringify(err));
                        });

                        //$scope is not updating so force with this command
                        if (!$scope.$$phase) { $scope.$apply(); }
                        scope.handleSave({ value: scope.todo });
                    };
                }
            };
         }
    }
}]);