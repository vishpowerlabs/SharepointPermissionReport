"use strict";
self["webpackHotUpdate_76d5422e_a579_4a00_99b3_192667104868_0_0_1"]("permission-viewer-web-part",{

/***/ 89135:
/*!**********************************************************************!*\
  !*** ./lib/webparts/permissionViewer/components/PermissionViewer.js ***!
  \**********************************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_14__ = __webpack_require__(/*! tslib */ 10196);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ 85959);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _fluentui_react_lib_Nav__WEBPACK_IMPORTED_MODULE_16__ = __webpack_require__(/*! @fluentui/react/lib/Nav */ 46786);
/* harmony import */ var _fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_17__ = __webpack_require__(/*! @fluentui/react/lib/Button */ 5613);
/* harmony import */ var _fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_25__ = __webpack_require__(/*! @fluentui/react/lib/Button */ 29425);
/* harmony import */ var _fluentui_react_lib_Spinner__WEBPACK_IMPORTED_MODULE_22__ = __webpack_require__(/*! @fluentui/react/lib/Spinner */ 80954);
/* harmony import */ var _fluentui_react_lib_Spinner__WEBPACK_IMPORTED_MODULE_23__ = __webpack_require__(/*! @fluentui/react/lib/Spinner */ 49885);
/* harmony import */ var _services_PermissionService__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../services/PermissionService */ 94110);
/* harmony import */ var _services_MockPermissionService__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../services/MockPermissionService */ 74632);
/* harmony import */ var _Header__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./Header */ 26783);
/* harmony import */ var _StatsCards__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./StatsCards */ 54088);
/* harmony import */ var _SitePermissions__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./SitePermissions */ 82561);
/* harmony import */ var _ListPermissions__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./ListPermissions */ 20548);
/* harmony import */ var _LoadingState__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ./LoadingState */ 37087);
/* harmony import */ var _DeepScanDialog__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ./DeepScanDialog */ 34937);
/* harmony import */ var _CheckAccess__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ./CheckAccess */ 69670);
/* harmony import */ var _SiteAdmins__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ./SiteAdmins */ 7021);
/* harmony import */ var _SiteGroups__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ./SiteGroups */ 55585);
/* harmony import */ var _SecurityGovernance__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! ./SecurityGovernance */ 48160);
/* harmony import */ var _utils_CsvExport__WEBPACK_IMPORTED_MODULE_15__ = __webpack_require__(/*! ../utils/CsvExport */ 2759);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_18__ = __webpack_require__(/*! @fluentui/react */ 63208);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_19__ = __webpack_require__(/*! @fluentui/react */ 46643);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_20__ = __webpack_require__(/*! @fluentui/react */ 4312);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_21__ = __webpack_require__(/*! @fluentui/react */ 10548);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_24__ = __webpack_require__(/*! @fluentui/react */ 87295);
/* harmony import */ var _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_13__ = __webpack_require__(/*! ./PermissionViewer.module.scss */ 29929);




















var PermissionViewer = function (props) {
    var _a;
    var _b = react__WEBPACK_IMPORTED_MODULE_0__.useState('site'), activeTab = _b[0], setActiveTab = _b[1];
    var _c = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), sitePermissions = _c[0], setSitePermissions = _c[1];
    var _d = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), filteredSitePermissions = _d[0], setFilteredSitePermissions = _d[1];
    var _e = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), lists = _e[0], setLists = _e[1];
    var _f = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), filteredLists = _f[0], setFilteredLists = _f[1];
    var _g = react__WEBPACK_IMPORTED_MODULE_0__.useState({ totalUsers: 0, totalGroups: 0, uniquePermissionsCount: 0 }), stats = _g[0], setStats = _g[1];
    var _h = react__WEBPACK_IMPORTED_MODULE_0__.useState(true), isLoading = _h[0], setIsLoading = _h[1];
    var _j = react__WEBPACK_IMPORTED_MODULE_0__.useState('Loading site permissions...'), loadingMessage = _j[0], setLoadingMessage = _j[1];
    var _k = react__WEBPACK_IMPORTED_MODULE_0__.useState(), permissionService = _k[0], setPermissionService = _k[1];
    var _l = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), isExporting = _l[0], setIsExporting = _l[1];
    // Site Admins State
    var _m = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), siteAdmins = _m[0], setSiteAdmins = _m[1];
    var _o = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), isLoadingAdmins = _o[0], setIsLoadingAdmins = _o[1];
    // Site Groups State
    var _p = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), siteGroups = _p[0], setSiteGroups = _p[1];
    var _q = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), isLoadingGroups = _q[0], setIsLoadingGroups = _q[1];
    var _r = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), isScanning = _r[0], setIsScanning = _r[1];
    // Deep Scan State
    var _s = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), isDeepScanOpen = _s[0], setIsDeepScanOpen = _s[1];
    var _t = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), deepScanItems = _t[0], setDeepScanItems = _t[1];
    var _u = react__WEBPACK_IMPORTED_MODULE_0__.useState(''), deepScanListTitle = _u[0], setDeepScanListTitle = _u[1];
    var _v = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), confirmScanList = _v[0], setConfirmScanList = _v[1];
    var _w = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), scanNoResults = _w[0], setScanNoResults = _w[1];
    var _x = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), errorMessage = _x[0], setErrorMessage = _x[1];
    // Delete Confirmation State
    var _y = react__WEBPACK_IMPORTED_MODULE_0__.useState({ isOpen: false, title: '', subText: '', onConfirm: function () { } }), deleteConfirmState = _y[0], setDeleteConfirmState = _y[1];
    // Access Control State
    var _z = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), hasAccess = _z[0], setHasAccess = _z[1]; // null = checking
    var _0 = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), accessContacts = _0[0], setAccessContacts = _0[1];
    react__WEBPACK_IMPORTED_MODULE_0__.useEffect(function () {
        var service;
        if (props.useMockData) {
            console.log("Using Mock Data");
            service = new _services_MockPermissionService__WEBPACK_IMPORTED_MODULE_2__.MockPermissionService();
        }
        else {
            service = new _services_PermissionService__WEBPACK_IMPORTED_MODULE_1__.PermissionServiceImpl(props.spHttpClient, props.webUrl);
        }
        setPermissionService(service);
        checkAccessAndLoad(service);
    }, [props.excludedLists, props.simulateAccessDenied, props.useMockData]);
    var checkAccessAndLoad = function (service) { return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__awaiter)(void 0, void 0, void 0, function () {
        var currentUser_1, isAllowed, owners, isOwner, admins, owners, mergedContacts_1, addContact, uniqueContacts, error_1;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    setIsLoading(true);
                    setLoadingMessage('Checking permissions...');
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 9, , 10]);
                    return [4 /*yield*/, service.getCurrentUser()];
                case 2:
                    currentUser_1 = _a.sent();
                    isAllowed = currentUser_1.IsSiteAdmin;
                    if (!!isAllowed) return [3 /*break*/, 4];
                    return [4 /*yield*/, service.getSiteOwners()];
                case 3:
                    owners = _a.sent();
                    isOwner = owners.some(function (o) { return o.LoginName === currentUser_1.LoginName || o.Email === currentUser_1.Email; });
                    if (isOwner) {
                        isAllowed = true;
                    }
                    _a.label = 4;
                case 4:
                    // SIMULATION OVERRIDE
                    if (props.simulateAccessDenied) {
                        console.log("Simulating Access Denied via Web Part Property");
                        isAllowed = false;
                    }
                    if (!isAllowed) return [3 /*break*/, 5];
                    setHasAccess(true);
                    // eslint-disable-next-line @typescript-eslint/no-floating-promises
                    loadData(service);
                    return [3 /*break*/, 8];
                case 5:
                    // Access Denied: Load Admins and Owners for contact info
                    setLoadingMessage('Loading contact information...');
                    return [4 /*yield*/, service.getSiteAdmins()];
                case 6:
                    admins = _a.sent();
                    return [4 /*yield*/, service.getSiteOwners()];
                case 7:
                    owners = _a.sent();
                    mergedContacts_1 = new Map();
                    addContact = function (u) {
                        var key = u.LoginName || u.Email || u.Title;
                        if (mergedContacts_1.has(key)) {
                            var existing = mergedContacts_1.get(key);
                            // Merge flags
                            existing.IsSiteAdmin = existing.IsSiteAdmin || u.IsSiteAdmin;
                            existing.IsSiteOwner = existing.IsSiteOwner || u.IsSiteOwner;
                        }
                        else {
                            mergedContacts_1.set(key, (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__assign)({}, u));
                        }
                    };
                    admins.forEach(addContact);
                    owners.forEach(addContact);
                    uniqueContacts = Array.from(mergedContacts_1.values());
                    setAccessContacts(uniqueContacts);
                    setHasAccess(false);
                    setIsLoading(false);
                    _a.label = 8;
                case 8: return [3 /*break*/, 10];
                case 9:
                    error_1 = _a.sent();
                    console.error("Error checking access", error_1);
                    // Default to denied on error for security
                    setHasAccess(false);
                    setIsLoading(false);
                    return [3 /*break*/, 10];
                case 10: return [2 /*return*/];
            }
        });
    }); };
    var loadData = function (service) { return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__awaiter)(void 0, void 0, void 0, function () {
        var sitePerms, siteStats, listsData, uniqueListsCount, admins, e_1, groups, e_2;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    // isLoading is already true from checkAccessAndLoad
                    setLoadingMessage('Loading site permissions...');
                    return [4 /*yield*/, service.getSiteRoleAssignments()];
                case 1:
                    sitePerms = _a.sent();
                    setSitePermissions(sitePerms);
                    setFilteredSitePermissions(sitePerms);
                    return [4 /*yield*/, service.getSiteStats()];
                case 2:
                    siteStats = _a.sent();
                    // Load lists
                    setLoadingMessage('Loading lists and libraries...');
                    return [4 /*yield*/, service.getLists(props.excludedLists)];
                case 3:
                    listsData = _a.sent();
                    uniqueListsCount = listsData.filter(function (l) { return l.HasUniqueRoleAssignments; }).length;
                    setStats((0,tslib__WEBPACK_IMPORTED_MODULE_14__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_14__.__assign)({}, siteStats), { uniquePermissionsCount: uniqueListsCount + (sitePerms.length > 0 ? 1 : 0) }));
                    // Load Site Admins
                    setIsLoadingAdmins(true);
                    _a.label = 4;
                case 4:
                    _a.trys.push([4, 6, 7, 8]);
                    return [4 /*yield*/, service.getSiteAdmins()];
                case 5:
                    admins = _a.sent();
                    setSiteAdmins(admins);
                    return [3 /*break*/, 8];
                case 6:
                    e_1 = _a.sent();
                    console.error("Error loading admins", e_1);
                    return [3 /*break*/, 8];
                case 7:
                    setIsLoadingAdmins(false);
                    return [7 /*endfinally*/];
                case 8:
                    // Load Site Groups
                    setIsLoadingGroups(true);
                    _a.label = 9;
                case 9:
                    _a.trys.push([9, 11, 12, 13]);
                    return [4 /*yield*/, service.getSiteGroups()];
                case 10:
                    groups = _a.sent();
                    setSiteGroups(groups);
                    return [3 /*break*/, 13];
                case 11:
                    e_2 = _a.sent();
                    console.error("Error loading groups", e_2);
                    return [3 /*break*/, 13];
                case 12:
                    setIsLoadingGroups(false);
                    return [7 /*endfinally*/];
                case 13:
                    // Sort lists: Unique first, then Inherited
                    listsData.sort(function (a, b) {
                        if (a.HasUniqueRoleAssignments === b.HasUniqueRoleAssignments)
                            return 0;
                        return a.HasUniqueRoleAssignments ? -1 : 1;
                    });
                    setLists(listsData);
                    setFilteredLists(listsData);
                    setIsLoading(false);
                    return [2 /*return*/];
            }
        });
    }); };
    var handleRefresh = function () {
        if (permissionService) {
            // eslint-disable-next-line @typescript-eslint/no-floating-promises
            loadData(permissionService);
        }
    };
    // ... (existing handlers)
    if (hasAccess === false) {
        return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_13__["default"].permissionViewer },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_13__["default"].webpartContainer },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_Header__WEBPACK_IMPORTED_MODULE_3__.Header, { onRefresh: handleRefresh, isLoading: isLoading, themeVariant: props.themeVariant, opacity: (_a = props.headerOpacity) !== null && _a !== void 0 ? _a : 100, title: props.webPartTitle, titleFontSize: props.webPartTitleFontSize }),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { style: { padding: '20px', textAlign: 'center' } },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("h2", { style: { color: '#d13438' } }, "You don't have access to this report"),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("p", { style: { marginBottom: '20px' } }, "Please check with the Site Owners or Site Administrators listed below."),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_SiteAdmins__WEBPACK_IMPORTED_MODULE_10__.SiteAdmins, { users: accessContacts, isLoading: false })))));
    }
    /* onSearch removed as unused */
    var handleGetListPermissions = function (listId) { return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__awaiter)(void 0, void 0, void 0, function () {
        var list;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!permissionService)
                        return [2 /*return*/, []];
                    list = lists.find(function (l) { return l.Id === listId; });
                    return [4 /*yield*/, permissionService.getListRoleAssignments(listId, list ? list.Title : '')];
                case 1: return [2 /*return*/, _a.sent()];
            }
        });
    }); };
    var handleRemoveSitePermission = function (principalId, principalName) {
        setDeleteConfirmState({
            isOpen: true,
            title: "Remove Permissions?",
            subText: "Are you sure you want to remove permissions for ".concat(principalName || 'this user', "? This will remove all permissions for this user on this site."),
            onConfirm: function () { void executeRemoveSitePermission(principalId); },
            onCancel: function () { }
        });
    };
    var executeRemoveSitePermission = function (principalId) { return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__awaiter)(void 0, void 0, void 0, function () {
        var success, sitePerms, error_2;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!permissionService)
                        return [2 /*return*/];
                    setDeleteConfirmState(function (prev) { return ((0,tslib__WEBPACK_IMPORTED_MODULE_14__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_14__.__assign)({}, prev), { isOpen: false })); });
                    setIsLoading(true);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 6, 7, 8]);
                    return [4 /*yield*/, permissionService.removeSitePermission(principalId)];
                case 2:
                    success = _a.sent();
                    if (!success) return [3 /*break*/, 4];
                    return [4 /*yield*/, permissionService.getSiteRoleAssignments()];
                case 3:
                    sitePerms = _a.sent();
                    setSitePermissions(sitePerms);
                    setFilteredSitePermissions(sitePerms);
                    return [3 /*break*/, 5];
                case 4:
                    setErrorMessage("Failed to remove permission. Please try again or check console for details.");
                    _a.label = 5;
                case 5: return [3 /*break*/, 8];
                case 6:
                    error_2 = _a.sent();
                    console.error("Error removing site permission", error_2);
                    setErrorMessage("An unexpected error occurred while removing permission.");
                    return [3 /*break*/, 8];
                case 7:
                    setIsLoading(false);
                    return [7 /*endfinally*/];
                case 8: return [2 /*return*/];
            }
        });
    }); };
    var handleRemoveListPermission = function (listId, principalId, principalName) { return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__awaiter)(void 0, void 0, void 0, function () {
        return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__generator)(this, function (_a) {
            return [2 /*return*/, new Promise(function (resolve) {
                    var onConfirmAction = function () { return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__awaiter)(void 0, void 0, void 0, function () {
                        var success, error_3;
                        return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__generator)(this, function (_a) {
                            switch (_a.label) {
                                case 0:
                                    setDeleteConfirmState(function (prev) { return ((0,tslib__WEBPACK_IMPORTED_MODULE_14__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_14__.__assign)({}, prev), { isOpen: false })); });
                                    if (!permissionService) {
                                        resolve(false);
                                        return [2 /*return*/];
                                    }
                                    _a.label = 1;
                                case 1:
                                    _a.trys.push([1, 3, , 4]);
                                    return [4 /*yield*/, permissionService.removeListPermission(listId, principalId)];
                                case 2:
                                    success = _a.sent();
                                    if (!success) {
                                        setErrorMessage("Failed to remove list permission. Please try again or check console for details.");
                                    }
                                    resolve(success);
                                    return [3 /*break*/, 4];
                                case 3:
                                    error_3 = _a.sent();
                                    console.error("Error removing list permission", error_3);
                                    setErrorMessage("An unexpected error occurred while removing list permission.");
                                    resolve(false);
                                    return [3 /*break*/, 4];
                                case 4: return [2 /*return*/];
                            }
                        });
                    }); };
                    setDeleteConfirmState({
                        isOpen: true,
                        title: "Remove Permissions?",
                        subText: "Are you sure you want to remove permissions for ".concat(principalName || 'this user', " on this list?"),
                        onConfirm: function () { void onConfirmAction(); },
                        onCancel: function () {
                            resolve(false);
                        }
                    });
                })];
        });
    }); };
    var handleRemoveDeepScanItemPermission = function (itemId, principalId, principalName) {
        setDeleteConfirmState({
            isOpen: true,
            title: "Remove Permissions?",
            subText: "Are you sure you want to remove permissions for ".concat(principalName || 'this user', " on this item?"),
            onConfirm: function () { void executeRemoveDeepScanItemPermission(itemId, principalId); },
            onCancel: function () { }
        });
    };
    var executeRemoveDeepScanItemPermission = function (itemId, principalId) { return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__awaiter)(void 0, void 0, void 0, function () {
        var list, success, error_4;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!permissionService || !deepScanListTitle)
                        return [2 /*return*/];
                    setDeleteConfirmState(function (prev) { return ((0,tslib__WEBPACK_IMPORTED_MODULE_14__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_14__.__assign)({}, prev), { isOpen: false })); });
                    list = lists.find(function (l) { return l.Title === deepScanListTitle; });
                    if (!list)
                        return [2 /*return*/];
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, permissionService.removeItemPermission(list.Id, itemId, principalId)];
                case 2:
                    success = _a.sent();
                    if (success) {
                        setDeepScanItems(function (prevItems) {
                            return prevItems.map(function (item) {
                                if (item.Id === itemId) {
                                    var newRoles = item.RoleAssignments.filter(function (ra) { return ra.PrincipalId !== principalId; });
                                    // Even if empty, we return the item with empty roles
                                    return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_14__.__assign)({}, item), { RoleAssignments: newRoles });
                                }
                                return item;
                            });
                        });
                    }
                    else {
                        setErrorMessage("Failed to remove item permission.");
                    }
                    return [3 /*break*/, 4];
                case 3:
                    error_4 = _a.sent();
                    console.error("Error removing item permission", error_4);
                    setErrorMessage("An unexpected error occurred while removing item permission.");
                    return [3 /*break*/, 4];
                case 4: return [2 /*return*/];
            }
        });
    }); };
    var handleRemoveFromGroup = function (groupId, userId, userName) {
        setDeleteConfirmState({
            isOpen: true,
            title: "Remove from Group?",
            subText: "Are you sure you want to remove '".concat(userName, "' from this group?"),
            onConfirm: function () { void executeRemoveFromGroup(groupId, userId); },
            onCancel: function () { }
        });
    };
    var executeRemoveFromGroup = function (groupId, userId) { return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__awaiter)(void 0, void 0, void 0, function () {
        var success, error_5;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!permissionService)
                        return [2 /*return*/];
                    setDeleteConfirmState(function (prev) { return ((0,tslib__WEBPACK_IMPORTED_MODULE_14__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_14__.__assign)({}, prev), { isOpen: false })); });
                    setIsLoading(true);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 6, 7, 8]);
                    return [4 /*yield*/, permissionService.removeUserFromGroup(groupId, userId)];
                case 2:
                    success = _a.sent();
                    if (!success) return [3 /*break*/, 4];
                    // We need to refresh the Site Permissions view if a group was expanded. 
                    // Since SitePermissions manages its own 'groupMembers' state but we don't have access to it easily here...
                    // Ideally, we should lift that state up or force a refresh.
                    // Re-loading the permissions effectively resets the view which is acceptable but collapses groups.
                    // A better UX would be to just reload the data.
                    // Reload data to ensure consistency
                    return [4 /*yield*/, loadData(permissionService)];
                case 3:
                    // We need to refresh the Site Permissions view if a group was expanded. 
                    // Since SitePermissions manages its own 'groupMembers' state but we don't have access to it easily here...
                    // Ideally, we should lift that state up or force a refresh.
                    // Re-loading the permissions effectively resets the view which is acceptable but collapses groups.
                    // A better UX would be to just reload the data.
                    // Reload data to ensure consistency
                    _a.sent();
                    return [3 /*break*/, 5];
                case 4:
                    setErrorMessage("Failed to remove user from group.");
                    _a.label = 5;
                case 5: return [3 /*break*/, 8];
                case 6:
                    error_5 = _a.sent();
                    console.error("Error removing user from group", error_5);
                    setErrorMessage("An unexpected error occurred.");
                    return [3 /*break*/, 8];
                case 7:
                    setIsLoading(false);
                    return [7 /*endfinally*/];
                case 8: return [2 /*return*/];
            }
        });
    }); };
    var handleExport = function () { return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__awaiter)(void 0, void 0, void 0, function () {
        var error_6;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!permissionService)
                        return [2 /*return*/];
                    setIsExporting(true);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 6, 7, 8]);
                    if (!(activeTab === 'site')) return [3 /*break*/, 3];
                    return [4 /*yield*/, (0,_utils_CsvExport__WEBPACK_IMPORTED_MODULE_15__.exportSitePermissions)(filteredSitePermissions, permissionService)];
                case 2:
                    _a.sent();
                    return [3 /*break*/, 5];
                case 3: return [4 /*yield*/, (0,_utils_CsvExport__WEBPACK_IMPORTED_MODULE_15__.exportListPermissions)(filteredLists, permissionService)];
                case 4:
                    _a.sent();
                    _a.label = 5;
                case 5: return [3 /*break*/, 8];
                case 6:
                    error_6 = _a.sent();
                    console.error("Export failed", error_6);
                    alert("Export failed. See console for details.");
                    return [3 /*break*/, 8];
                case 7:
                    setIsExporting(false);
                    return [7 /*endfinally*/];
                case 8: return [2 /*return*/];
            }
        });
    }); };
    // ...
    var handleDeepScan = function (listId) { return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__awaiter)(void 0, void 0, void 0, function () {
        var list;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__generator)(this, function (_a) {
            list = lists.find(function (l) { return l.Id === listId; });
            if (!list)
                return [2 /*return*/];
            setScanNoResults(false);
            setConfirmScanList({ id: list.Id, title: list.Title });
            return [2 /*return*/];
        });
    }); };
    var executeDeepScan = function () { return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__awaiter)(void 0, void 0, void 0, function () {
        var listId, listTitle, items, e_3;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_14__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!permissionService || !confirmScanList)
                        return [2 /*return*/];
                    listId = confirmScanList.id;
                    listTitle = confirmScanList.title;
                    setIsScanning(true);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, 4, 5]);
                    return [4 /*yield*/, permissionService.getUniquePermissionItems(listId)];
                case 2:
                    items = _a.sent();
                    if (items.length > 0) {
                        setDeepScanItems(items);
                        setDeepScanListTitle(listTitle);
                        setIsDeepScanOpen(true);
                        setConfirmScanList(null); // Close confirm dialog only if opening results
                    }
                    else {
                        setScanNoResults(true);
                        // Keep dialog open to show message
                    }
                    return [3 /*break*/, 5];
                case 3:
                    e_3 = _a.sent();
                    console.error(e_3);
                    alert("Error during deep scan.");
                    setConfirmScanList(null); // Close on error too
                    return [3 /*break*/, 5];
                case 4:
                    setIsScanning(false);
                    return [7 /*endfinally*/];
                case 5: return [2 /*return*/];
            }
        });
    }); };
    var downloadDeepScanResults = function () {
        (0,_utils_CsvExport__WEBPACK_IMPORTED_MODULE_15__.exportDeepScanResults)(deepScanItems, deepScanListTitle);
    };
    var getDialogTitle = function () {
        if (scanNoResults)
            return 'Deep Scan Complete';
        if (isScanning)
            return 'Deep Scan in Progress...';
        return 'Start Deep Scan?';
    };
    var getDialogSubText = function () {
        if (scanNoResults) {
            return "No items with unique permissions were found in \"".concat(confirmScanList === null || confirmScanList === void 0 ? void 0 : confirmScanList.title, "\". All items inherit permissions.");
        }
        if (isScanning) {
            return "Scanning \"".concat(confirmScanList === null || confirmScanList === void 0 ? void 0 : confirmScanList.title, "\". This may take a few moments depending on the number of items...");
        }
        return "This will verify every single item in \"".concat(confirmScanList === null || confirmScanList === void 0 ? void 0 : confirmScanList.title, "\" to find unique permissions. This might take a while for large lists. Continue?");
    };
    var navGroups = [
        {
            links: [
                {
                    name: 'Site Permissions',
                    url: '',
                    key: 'site',
                    icon: 'Shield',
                    onClick: function () { setActiveTab('site'); }
                },
                {
                    name: 'Lists & Libraries',
                    url: '',
                    key: 'lists',
                    icon: 'List',
                    onClick: function () { setActiveTab('lists'); }
                },
                {
                    name: 'Security & Governance',
                    url: '',
                    key: 'governance',
                    icon: 'SecurityGroup',
                    onClick: function () { setActiveTab('governance'); }
                },
                {
                    name: 'Site Groups',
                    url: '',
                    key: 'groups',
                    icon: 'Group',
                    onClick: function () { setActiveTab('groups'); }
                },
                {
                    name: 'Check Access',
                    url: '',
                    key: 'check_access',
                    icon: 'UserOptional',
                    onClick: function () { setActiveTab('check_access'); }
                },
                {
                    name: 'Site Admins',
                    url: '',
                    key: 'admins',
                    icon: 'Admin',
                    onClick: function () { setActiveTab('admins'); }
                }
            ].filter(function (link) { return link.key !== 'governance' || props.showSecurityGovernanceTab !== false; })
        }
    ];
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_13__["default"].permissionViewer },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_13__["default"].webpartContainer },
            (props.showComponentHeader !== false) && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_Header__WEBPACK_IMPORTED_MODULE_3__.Header, { title: props.webPartTitle, titleFontSize: props.webPartTitleFontSize, themeVariant: props.themeVariant, opacity: props.headerOpacity, isLoading: isLoading, onRefresh: handleRefresh })),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_13__["default"].layoutContainer },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_13__["default"].navigation },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Nav__WEBPACK_IMPORTED_MODULE_16__.Nav, { groups: navGroups, selectedKey: activeTab, styles: {
                            root: {
                                width: '100%',
                                height: '100%',
                                boxSizing: 'border-box',
                                border: '1px solid transparent',
                                overflowY: 'auto'
                            }
                        } })),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_13__["default"].mainCanvas },
                    (props.showStats !== false) && react__WEBPACK_IMPORTED_MODULE_0__.createElement(_StatsCards__WEBPACK_IMPORTED_MODULE_4__.StatsCards, { stats: stats, highlight: false }),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_13__["default"].toolbar }, (activeTab === 'site' || activeTab === 'lists') && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_17__.DefaultButton, { text: isExporting ? "Exporting..." : "Export to CSV", iconProps: { iconName: 'Download' }, onClick: function () { void handleExport(); }, disabled: isExporting || isLoading, className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_13__["default"].exportBtn, styles: {
                            root: { height: '32px' }, // Maintain height
                            label: { fontSize: props.buttonFontSize || '12px', fontWeight: 600 }
                        } }))),
                    errorMessage && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_18__.MessageBar, { messageBarType: _fluentui_react__WEBPACK_IMPORTED_MODULE_19__.MessageBarType.error, isMultiline: false, onDismiss: function () { return setErrorMessage(null); }, dismissButtonAriaLabel: "Close", styles: { root: { marginBottom: 10 } } }, errorMessage)),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_13__["default"].content },
                        isLoading && react__WEBPACK_IMPORTED_MODULE_0__.createElement(_LoadingState__WEBPACK_IMPORTED_MODULE_7__.LoadingState, { message: loadingMessage }),
                        !isLoading && activeTab === 'site' && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_SitePermissions__WEBPACK_IMPORTED_MODULE_5__.SitePermissions, { permissions: filteredSitePermissions, permissionService: permissionService, contentFontSize: props.contentFontSize, onRemovePermission: handleRemoveSitePermission, onRemoveFromGroup: handleRemoveFromGroup })),
                        !isLoading && activeTab === 'lists' && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_ListPermissions__WEBPACK_IMPORTED_MODULE_6__.ListPermissions, { lists: filteredLists, getListPermissions: handleGetListPermissions, onScanItems: handleDeepScan, themeVariant: props.themeVariant, buttonFontSize: props.buttonFontSize, contentFontSize: props.contentFontSize, onRemovePermission: handleRemoveListPermission, forcedExpandedListId: null })),
                        !isLoading && activeTab === 'groups' && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_SiteGroups__WEBPACK_IMPORTED_MODULE_11__.SiteGroups, { groups: siteGroups, isLoading: isLoadingGroups, permissionService: permissionService, contentFontSize: props.contentFontSize })),
                        !isLoading && activeTab === 'check_access' && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_CheckAccess__WEBPACK_IMPORTED_MODULE_9__.CheckAccess, { permissionService: permissionService, sitePermissions: sitePermissions, lists: lists, contentFontSize: props.contentFontSize })),
                        !isLoading && activeTab === 'governance' && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_SecurityGovernance__WEBPACK_IMPORTED_MODULE_12__.SecurityGovernance, { contentFontSize: props.contentFontSize, permissionService: permissionService, showExternalUserAudit: props.showExternalUserAudit, showSharingLinks: props.showSharingLinks, showOrphanedUsers: props.showOrphanedUsers })),
                        !isLoading && activeTab === 'admins' && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_SiteAdmins__WEBPACK_IMPORTED_MODULE_10__.SiteAdmins, { users: siteAdmins, isLoading: isLoadingAdmins })))),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_DeepScanDialog__WEBPACK_IMPORTED_MODULE_8__.DeepScanDialog, { isOpen: isDeepScanOpen, onDismiss: function () { return setIsDeepScanOpen(false); }, listTitle: deepScanListTitle, items: deepScanItems, onDownload: downloadDeepScanResults, buttonFontSize: props.buttonFontSize, contentFontSize: props.contentFontSize, onRemovePermission: handleRemoveDeepScanItemPermission }),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_20__.Dialog, { hidden: !confirmScanList, onDismiss: function () { if (!isScanning)
                        setConfirmScanList(null); }, dialogContentProps: {
                        type: _fluentui_react__WEBPACK_IMPORTED_MODULE_21__.DialogType.normal,
                        title: getDialogTitle(),
                        subText: getDialogSubText()
                    } }, isScanning ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Spinner__WEBPACK_IMPORTED_MODULE_22__.Spinner, { size: _fluentui_react_lib_Spinner__WEBPACK_IMPORTED_MODULE_23__.SpinnerSize.large, label: "Scanning for unique permissions..." })) : (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_24__.DialogFooter, null, scanNoResults ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_25__.PrimaryButton, { onClick: function () { return setConfirmScanList(null); }, text: "Close", styles: {
                        root: { height: '32px' },
                        label: { fontSize: props.buttonFontSize || '12px', fontWeight: 600 }
                    } })) : (react__WEBPACK_IMPORTED_MODULE_0__.createElement(react__WEBPACK_IMPORTED_MODULE_0__.Fragment, null,
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_25__.PrimaryButton, { onClick: executeDeepScan, text: "Start Scan", styles: {
                            root: { height: '32px' },
                            label: { fontSize: props.buttonFontSize || '12px', fontWeight: 600 }
                        } }),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_17__.DefaultButton, { onClick: function () { return setConfirmScanList(null); }, text: "Cancel", styles: {
                            root: { height: '32px' },
                            label: { fontSize: props.buttonFontSize || '12px', fontWeight: 600 }
                        } })))))),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_20__.Dialog, { hidden: !deleteConfirmState.isOpen, onDismiss: function () { return setDeleteConfirmState(function (prev) { return ((0,tslib__WEBPACK_IMPORTED_MODULE_14__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_14__.__assign)({}, prev), { isOpen: false })); }); }, dialogContentProps: {
                        type: _fluentui_react__WEBPACK_IMPORTED_MODULE_21__.DialogType.normal,
                        title: deleteConfirmState.title,
                        subText: deleteConfirmState.subText,
                    }, modalProps: {
                        isBlocking: true,
                        styles: { main: { maxWidth: 450 } }
                    } },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_24__.DialogFooter, null,
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_25__.PrimaryButton, { onClick: deleteConfirmState.onConfirm, text: "Remove", styles: {
                                root: { background: '#d13438', border: '1px solid #d13438' }, // Red color for danger
                                rootHovered: { background: '#a4262c' },
                                label: { fontWeight: 600 }
                            } }),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_17__.DefaultButton, { onClick: function () { return setDeleteConfirmState(function (prev) { return ((0,tslib__WEBPACK_IMPORTED_MODULE_14__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_14__.__assign)({}, prev), { isOpen: false })); }); }, text: "Cancel" })))))));
};
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (PermissionViewer);


/***/ })

},
/******/ function(__webpack_require__) { // webpackRuntimeModules
/******/ /* webpack/runtime/getFullHash */
/******/ (() => {
/******/ 	__webpack_require__.h = () => ("cca8d54a935346eeefd6")
/******/ })();
/******/ 
/******/ }
);
//# sourceMappingURL=permission-viewer-web-part.91ff3b5a2e4d3b5a2a33.hot-update.js.map