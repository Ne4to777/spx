// variable to hold the following status of the page
let pageIsFollowed

// Get the url of the current page
function getUrl() {
	const ctxCurrent = SP.ClientContext.get_current()
	const url = ctxCurrent.get_url()
	const docUrl = `${window.location.protocol}//${window.location.host}${url}?listId=${_spPageContextInfo.pageListId}&itemId=${_spPageContextInfo.pageItemId}`
	return docUrl
}

// Check if the current page is already followed
function isAlreadyFollowed() {
	const clientContext = SP.ClientContext.get_current()
	const socialManager = new SP.Social.SocialFollowingManager(clientContext)
	const socialActor = new SP.Social.SocialActorInfo()
	socialActor.set_contentUri(getUrl())
	socialActor.set_actorType(SP.Social.SocialActorTypes.documents)

	this.result = socialManager.isFollowed(socialActor)
	clientContext.executeQueryAsync(Function.createDelegate(this, this.onCheckFollowSucceeded), Function.createDelegate(this, this.onQueryFailed))
}

// Toggle the star if followed or not using jQuery
function onCheckFollowSucceeded() {
	$('#favoriteLink').toggleClass('favIconSelected', this.result.get_value())
	// set global variable
	pageIsFollowed = this.result.get_value()
}

// Alert error
function onQueryFailed(sender, args) {
	alert(`Request failed. ${args.get_message()}\n${args.get_stackTrace()}`)
}

// Follow the page
function startFollowing() {
	const clientContext = SP.ClientContext.get_current()
	const socialManager = new SP.Social.SocialFollowingManager(clientContext)
	const siteActorInfo = new SP.Social.SocialActorInfo()
	siteActorInfo.set_contentUri(getUrl())
	siteActorInfo.set_actorType(SP.Social.SocialActorTypes.documents)
	// follow call
	socialManager.follow(siteActorInfo)
	// upon success, we reexecute isAlreadyFollowed to make sure it was followed - a bit lazy :)
	clientContext.executeQueryAsync(Function.createDelegate(this, this.isAlreadyFollowed), Function.createDelegate(this, this.onQueryFailed))
}

// Stop follow the page
function stopFollowing() {
	const clientContext = SP.ClientContext.get_current()
	const socialManager = new SP.Social.SocialFollowingManager(clientContext)
	const siteActorInfo = new SP.Social.SocialActorInfo()
	siteActorInfo.set_contentUri(getUrl())
	siteActorInfo.set_actorType(SP.Social.SocialActorTypes.documents)
	// stop following call
	socialManager.stopFollowing(siteActorInfo)
	// upon success, we reexecute isAlreadyFollowed to make sure it was stopped - a bit lazy :)
	clientContext.executeQueryAsync(Function.createDelegate(this, this.isAlreadyFollowed), Function.createDelegate(this, this.onQueryFailed))
}

// function attached to the click event of the star
function toggleFollow() {
	if (!pageIsFollowed) {
		startFollowing()
	} else {
		stopFollowing()
	}
}

// on load, check if page is followed or not and toggle star accordingly
// $(document).ready(() => {
// 	// Make sure SP.ClientContext is available
// 	SP.SOD.executeFunc('sp.js', 'SP.ClientContext', () => {
// 		// Make sure SocialFollowingManager is available
// 		SP.SOD.executeFunc('userprofile', 'SP.Social.SocialFollowingManager', () => {
// 			isAlreadyFollowed()
// 		})
// 	})
// })
