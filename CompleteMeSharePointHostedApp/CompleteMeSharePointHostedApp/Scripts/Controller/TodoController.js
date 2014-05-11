myApp.controller('TodoCtrl', ['$scope', 'SharePointJSOMService', function ($scope, SharePointJSOMService) {
    SP.SOD.executeOrDelayUntilScriptLoaded(runMyCode, "SP.js");
    function runMyCode() {
        $scope.todos = [];

        $.when(SharePointJSOMService.getCurrentUser())
        .done(function (jsonObject) {
            currentUser = jsonObject.d;
        })
        .fail(function (err) {
            console.info(JSON.stringify(err));
        });

        $.when(SharePointJSOMService.checkTaskList($scope, TaskListName))
        .done(function (jsonObject) {
        })
        .fail(function (err) {
            console.info(JSON.stringify(err));
        });

        $.when(SharePointJSOMService.getTasksREST($scope, TaskListName))
        .done(function (jsonObject) {
            angular.forEach(jsonObject.d.results, function (todo) {
                $scope.todos.push({
                    text: todo.Title,
                    dueDate: todo.DueDate,
                    priority: todo.Priority,
                    assignedTo: todo.AssignedTo.Title,
                    assignedToName: todo.AssignedTo.Name,
                    status: todo.Status,
                    done: todo.Status == 'Completed',
                    id: todo.ID
                });
                //$scope is not updating so force with this command
                if (!$scope.$$phase) { $scope.$apply(); }
            });
        })
        .fail(function (err) {
            console.info(JSON.stringify(err));
        });

        //add function on controller
        $scope.addTask = function ($event) {
            $event.preventDefault();
            
            $.when(SharePointJSOMService.addTask(TaskListName, $scope.todoText, $scope.dt))
            .done(function (id) {
                $scope.todos.push({ text: $scope.todoText, dueDate: $scope.dt, done: false, id: id, assignedTo: currentUser.Title, assignedToName: currentUser.LoginName.split("|")[1] });

                //$scope is not updating so force with this command
                if (!$scope.$$phase) { $scope.$apply(); }

                $scope.todoText = '';
                $scope.dt = new Date();
            })
            .fail(function (err) {
                console.info(JSON.stringify(err));
            });

        };

        $scope.remaining = function () {
            var count = 0;
            angular.forEach($scope.todos, function (todo) {
                count += todo.done ? 0 : 1;
            });
            return count;
        };

        //datepicker
        $scope.today = function () {
            $scope.dt = new Date();
        };
        $scope.today();

        $scope.showWeeks = true;
        $scope.toggleWeeks = function () {
            $scope.showWeeks = !$scope.showWeeks;
        };

        $scope.toggleMin = function () {
            $scope.minDate = ($scope.minDate) ? null : new Date();
        };
        $scope.toggleMin();

        $scope.open = function ($event) {
            $event.preventDefault();
            $event.stopPropagation();

            $scope.opened = true;
        };

        $scope.dateOptions = {
            'year-format': "'yy'",
            'starting-day': 1
        };

        $scope.formats = ['dd-MMM-yyyy', 'yyyy/MM/dd', 'shortDate'];
        $scope.format = $scope.formats[0];
    }
}]);